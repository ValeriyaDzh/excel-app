import os
import pandas as pd
import smtplib
from dotenv import load_dotenv
from pathlib import Path

from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

from lexicon import MESSAGE, COLUMNS

load_dotenv()


def send_file_to_providers(file_name: str, to_email: str, text: str) -> None:
    """Функция для рассылки по поставщикам.
    file_name: имя отправляемого файла
    to_email: email поставщика
    text: текст в теле письма
    """
    sender = os.getenv("SENDER_EMAIL")
    password = os.getenv("PASSWORD")
    subject = Path(file_name).name
    body = text

    server = smtplib.SMTP("smtp.yandex.ru", 587)
    server.starttls()

    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body))

    with open(file_name, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("content-disposition", "attachment", filename=subject)
        msg.attach(part)

    try:
        server.login(sender, password)
        server.sendmail(sender, to_email, msg.as_string())
    except Exception as _ex:
        print(f'{_ex}{MESSAGE["sendmail_error"]}')
    finally:
        server.quit()


def split_file_to_providers(
    file_path: str, name_split: str, providers_email: str
) -> str:
    """Функция для разделения единного файла для рассылки поставщикам.
    file_path: путь к файлу формата .xlsx
    name_split: название итогового файла для поставщика
    providers_email: путь к файлу с почтой поставщиков
    """

    file_path = Path(file_path)
    providers_email_path = Path(providers_email)
    if file_path.suffix == ".xlsx":
        print(f'\n{file_path.name} {MESSAGE["at_work"]}')

        full_file = pd.read_excel(file_path.name)
        provider_email = pd.read_excel(providers_email_path.name)
        provider_data = full_file.groupby(COLUMNS["provider"])

        directory_path = file_path.parent / "providers_file"
        directory_path.mkdir(parents=True, exist_ok=True)

        created_files = []
        for provider, data in provider_data:
            file = directory_path / f"{provider[3:]}_{name_split}.xlsx"
            data.to_excel(file, index=False)
            created_files.append((provider, str(file)))

        for_send_data = pd.DataFrame(
            created_files, columns=(COLUMNS["provider"], COLUMNS["file"])
        )
        for_send_data = pd.merge(
            for_send_data, provider_email, how="left", on=COLUMNS["provider"]
        )

        send_confirmation = input(MESSAGE["dispatch_question"])
        if send_confirmation.lower() in MESSAGE["confirm_answers"]:
            mail_text = input(
                f'{MESSAGE["text_for_mail_question"]} "{MESSAGE["text_for_mail_default"]}"'
            )
            mail_text = mail_text if mail_text else MESSAGE["text_for_mail_default"]
            print(MESSAGE["process_sending"])
            for index, row in for_send_data.iterrows():
                send_file_to_providers(
                    row[COLUMNS["file"]], row[COLUMNS["email"]], mail_text
                )
            print(MESSAGE["success_sending"])

        return f'{MESSAGE["files_save_in"]} {directory_path}'
    else:
        return MESSAGE["type_error"]


def main():
    user_file_path = input(MESSAGE["file_path_question"])
    user_name_split = input(MESSAGE["name_split_question"])
    user_providers_path = input(MESSAGE["providers_path_question"])
    print(
        split_file_to_providers(
            file_path=user_file_path,
            name_split=user_name_split,
            providers_email=user_providers_path,
        )
    )


if __name__ == "__main__":
    main()

import pandas as pd
from pathlib import Path
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

from lexicon import MESSAGE


def send_file_to_providers(file_name: str, to_email: str, text: str) -> None:
    sender = ""
    password = ""
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
    file_path: str,
    name_split: str = "отборка_2024",
    providers_email: str = "Почта поставщиков.xlsx",
):
    """Функция для разделения единного файла для рассылки поставщикам.
    file_path: путь к файлу формата .xlsx
    name_split: название итогового файла для поставщика
    providers_email: путь к файлу с почтой поставщиков
    """

    if Path(file_path).suffix == ".xlsx":
        print(f'\n{Path(file_path).name} {MESSAGE["at_work"]}')

        full_file = pd.read_excel(Path(file_path).name)
        provider_email = pd.read_excel(providers_email)
        provider_data = full_file.groupby("Поставщик")
        #  удалить или оставить возможность сохранить итог
        # for_send_mail = Path(file_path).parent / f"файлы для отправки {name_split}.xlsx"

        directory_path = Path(file_path).parent / "providers_file"

        if not directory_path.exists():
            directory_path.mkdir()

        created_files = []
        for provider, data in provider_data:
            file = directory_path / f"{provider[3:]}_{name_split}.xlsx"
            data.to_excel(file, index=False)
            created_files.append((provider, str(file)))

        df = pd.DataFrame(created_files, columns=("Поставщик", "Файл"))
        for_send_data = pd.merge(df, provider_email, how="left", on="Поставщик")

        send_file = input(MESSAGE["dispatch_question"])
        if send_file.lower() in MESSAGE["confirm_answers"]:
            text = input(
                MESSAGE["text_for_mail_question"], MESSAGE["text_for_mail_default"]
            )
            if not text:
                text = MESSAGE["text_for_mail_default"]
            print(MESSAGE["process_sending"])
            for index, row in for_send_data.iterrows():
                send_file_to_providers(row["Файл"], row["Почта"], text)
            print(MESSAGE["success_sending"])

        return f'{MESSAGE["files_save_in"]} {directory_path}'
    else:
        return MESSAGE["type_error"]


def main():
    user_file_path = input(MESSAGE["file_path_question"])
    user_name_split = input(MESSAGE["name_split_question"])
    split_file_to_providers(file_path=user_file_path, name_split=user_name_split)


if __name__ == "__main__":
    main()

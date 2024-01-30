import pandas as pd
from pathlib import Path
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders


def send_file_to_providers(file_name, to_email, text):
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
        print(f"{_ex}\nПроверьте логин и пароль")
    finally:
        server.quit()


def split_file_to_providers(
    file_name="",
    name_split="отборка_2024",
    providers_email="Почта поставщиков.xlsx",
):
    """Функция для разделения единного файла для рассылки поставщикам.
    file_name - имя файла формата .xlsx
    name_split - название итогового файла для поставщика
    """

    if Path(file_name).suffix == ".xlsx":
        print(f"\n{Path(file_name).name} в работе")

        full_file = pd.read_excel(Path(file_name).name)
        provider_email = pd.read_excel(providers_email)
        provider_data = full_file.groupby("Поставщик")
        for_send_mail = Path.cwd() / f"файлы для отправки {name_split}.xlsx"

        created_files = []
        for provider, data in provider_data:
            file = Path.cwd() / "providers_file" / f"{provider[3:]}_{name_split}.xlsx"
            data.to_excel(file, index=False)
            created_files.append((provider, str(file)))

        df = pd.DataFrame(created_files, columns=("Поставщик", "Файл"))
        for_send_data = pd.merge(df, provider_email, how="left", on="Поставщик")
        # for_send_data.to_excel(for_send_mail, index=False)                        # IF NEED SAVE DATA AS .xlsx

        send_file = input("\nОтправить файлы поставщикам? да/нет\n")
        if send_file.lower() == "да":
            text = input(
                '\nВведите текст письма или нажмите enter\nтекст по-умолчанию:"Файл во вложении прошу ознакомиться":\n'
            )
            if not text:
                text = "Файл во вложении прошу ознакомиться"
            print("\nОтправка писем...")
            for index, row in for_send_data.iterrows():
                send_file_to_providers(row["Файл"], row["Почта"], text)
            print("Письма отправлены успешно")

        return f"\nФайлы сохранены: {Path.cwd() / 'providers_file'}"
    else:
        return f"\nДля обработки данных файл должен быть 'xlsx'"


def main():
    user_file_name = input("\nВведите название общего файла:\n")
    user_name_split = input("\nВведите название к файлу для поставщика:\n")
    split_file_to_providers(file_name=user_file_name, name_split=user_name_split)


if __name__ == "__main__":
    main()

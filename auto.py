import pandas as pd
import matplotlib.pyplot as plt
from smtplib import SMTP
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

def generate_report(data_file: str, email: str):
    try:
        # Чтение и обработка данных из Excel-файла
        df = pd.read_excel(data_file, engine='openpyxl')
        summary = df.describe()

        # Создание визуализации
        plt.figure(figsize=(10, 6))
        df.plot(kind='line')
        plt.title('Обзор данных о продажах')
        plt.xlabel('Дата')
        plt.ylabel('Продажи')
        plt.grid(True)
        plt.tight_layout()

        # Сохранение графика
        plot_file = 'sales_report.png'
        plt.savefig(plot_file)

        # Отправка отчета по электронной почте
        send_email(email, plot_file, summary.to_string())
    except Exception as e:
        print(f"Произошла ошибка: {e}")

def send_email(to_email: str, image_path: str, summary: str):
    msg = MIMEMultipart()
    msg['Subject'] = 'Автоматический отчет по продажам'
    msg['From'] = 'your_email@example.com'
    msg['To'] = to_email

    text = MIMEText(f"Добрый день,\n\nПожалуйста, найдите отчет по продажам ниже:\n\n{summary}", 'plain')
    msg.attach(text)

    with open(image_path, 'rb') as img:
        img_data = img.read()
        image = MIMEImage(img_data, name='Отчет о продажах')
        msg.attach(image)

    with SMTP('smtp.example.com', 587) as smtp:
        smtp.starttls()
        smtp.login('your_email@example.com', 'your_password')
        smtp.sendmail(msg['From'], msg['To'], msg.as_string())

# Пример использования
generate_report('sales_data.xlsx', 'recipient@example.com')

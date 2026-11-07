#Указываем базовый образ системы
FROM python:3.12-slim

#Указываем рабочую папку приложения, где будут находится файлы
WORKDIR /app

# Устанавливаем необходимые системные пакеты
RUN apt-get update && apt-get install -y --no-install-recommends \
  ca-certificates \
  wget

# Устанавливаем Python-библиотеки
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt


#Устанавливаем сертификаты минцифры
RUN wget -O /usr/local/share/ca-certificates/russian_trusted_root_ca.crt \
    https://gu-st.ru/content/lending/russian_trusted_root_ca_pem.crt \
    && wget -O /usr/local/share/ca-certificates/russian_trusted_sub_ca.crt \
    https://gu-st.ru/content/lending/russian_trusted_sub_ca_pem.crt

RUN update-ca-certificates 

# Копируем исходный код (bot.py что -> . текущая папка из WORKDIR)
COPY osnova.py template.pptx TOKEN.env ./

# Запускаем Python скрипт
CMD ["python", "-u", "osnova.py"]

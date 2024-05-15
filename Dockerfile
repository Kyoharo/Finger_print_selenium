FROM python:3.9

WORKDIR /app

COPY . .

RUN pip install -r requirements.txt

COPY geckodriver /app/geckodriver 
COPY chromedriver /app/chromedriver  


# Entrypoint command
ENTRYPOINT ["python", "mainb.py"]

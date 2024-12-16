FROM python:3.10 as base_builder

ENV TZ=Europe/Moscow

WORKDIR /app
ADD ./requirements.txt /app/requirements.txt

RUN python -m pip install -r ./requirements.txt

ADD ./ /app/
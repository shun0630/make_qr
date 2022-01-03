FROM ubuntu:latest
RUN apt-get update && apt-get install -y fontconfig
RUN apt-get install -y language-pack-ja-base language-pack-ja locales
RUN apt-get install -y python3 python3-pip
ENV locale-gen ja_JP.UTF-8
ENV TERM xterm
RUN apt -y install fonts-ipaexfont

RUN mkdir -p /usr/share/fonts
COPY ./fonts /usr/share/fonts

RUN mkdir /app

COPY ./requirements.txt ./

RUN pip install -r ./requirements.txt
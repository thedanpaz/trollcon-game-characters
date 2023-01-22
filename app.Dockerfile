# Dockerfile References: https://docs.docker.com/engine/reference/builder/

FROM python:3.10-buster

RUN pip3 install openpyxl

CMD tail -f /dev/null
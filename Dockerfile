FROM python:3.12-bookworm

WORKDIR /usr/src/app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY convert_to_matches.py .
COPY run.sh .

CMD [ "bash", "run.sh" ]

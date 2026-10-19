# Linux image for the test suite on a virtual display (tools/test_headless.sh): no window on your screen.
# Same as CI's Linux job: Python 3.13, Tk, Xvfb. The code (this repo and ../simple-core) is mounted at run time.
FROM python:3.13-bookworm
RUN apt-get update \
    && apt-get install -y --no-install-recommends xvfb xauth fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*
COPY requirements.txt requirements-dev.txt /tmp/req/
# (pyinstaller is for builds only; simplecore comes from the mounted ../simple-core)
RUN grep -v -i "^pyinstaller" /tmp/req/requirements-dev.txt > /tmp/req/test.txt \
    && cd /tmp/req && pip install --no-cache-dir -r test.txt
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONPATH=/core
WORKDIR /app

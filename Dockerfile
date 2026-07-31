FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py companies.py rcsaero_client.py sheet_builder.py Department_Template.xlsx Category_Template.xlsx Pricing_Template.xlsx ./
COPY templates/ templates/

ENV PORT=5000

EXPOSE 5000

CMD ["python", "app.py"]

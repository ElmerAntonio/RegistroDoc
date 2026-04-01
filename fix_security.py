with open("src/rdsecurity.py", "r") as f:
    content = f.read()

# Fix the datetime.utcnow() deprecation warning found during pytest
content = content.replace("datetime.datetime.utcnow()", "datetime.datetime.now(datetime.timezone.utc)")

with open("src/rdsecurity.py", "w") as f:
    f.write(content)

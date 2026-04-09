with open("src/rddata.py", "r") as f:
    content = f.read()

content = content.replace("should_close = False\n        if wb is None:\n            should_close = not bool(self._wb_cache)", "should_close = not bool(self._wb_cache) if wb is None else False")

with open("src/rddata.py", "w") as f:
    f.write(content)

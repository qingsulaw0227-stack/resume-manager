# Contributing

Contributions are welcome. Use invented names, organizations, dates and contact
details in tests and screenshots. Never submit a real person's resume or photo.

Before opening a pull request, run:

```sh
python -m unittest discover -s tests -v
```

Changes to Word generation should preserve editable DOCX output and add a test
for the affected structure. Layout claims must be based on an actual render;
successful file creation alone is not proof of correct pagination.

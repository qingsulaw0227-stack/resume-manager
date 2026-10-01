# Security and privacy

Resume Kitchen is designed for one person or one trusted household to run on a
private computer. It has no login system and must not be exposed directly to the
public internet.

## Safe deployment

- Keep `BIND_HOST=127.0.0.1`, or use the supplied Docker Compose mapping.
- Store `data/` outside public repositories and include it in backups.
- Do not commit `.env`, photographs, resumes, generated Word files or API keys.
- If remote access is needed, use an authenticated private tunnel and review its
  provider's privacy and pricing terms.
- Treat imported documents as untrusted. The importer reads text only and does
  not execute macros, embedded programs or document instructions.

## Reporting a vulnerability

Open a GitHub security advisory for the repository rather than including private
resume data in a public issue. Remove personal files, names and contact details
from reproduction cases.

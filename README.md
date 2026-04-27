# Gmail Promotions Classifier CLI

Python CLI for classifying Gmail messages, applying Gmail labels, and writing an audit log. It supports both OpenAI and Ollama backends and is now set up so you can publish the repository without committing personal credentials or local mail data.

## What It Does

- Connects to Gmail over IMAP using a Gmail App Password.
- Scans either a specific IMAP folder or a Gmail query.
- Extracts sender, subject, date, `Message-ID`, `List-Unsubscribe`, and a text snippet.
- Classifies messages with deterministic rules plus an LLM backend.
- Applies Gmail labels such as `AI/KEEP`, `AI/DELETE`, and `AI/PROCESSED`.
- Writes an audit CSV for review.

## Share-Safe Repo Setup

This repo is configured so local-only files stay out of git.

Ignored local files include:

- `.env.local` for credentials and API keys
- `allowlist.txt` for personal keep rules
- `audit_*.csv` and related audit exports
- Python cache and virtual environment folders

Tracked template files include:

- `.env.local.example`
- `allowlist.example.txt`

The CLI auto-loads `.env.local` before parsing arguments, and it now includes an interactive setup wizard for first-time users.

## Quick Start

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the guided setup wizard:

```bash
python main.py --setup
```

The wizard walks you through:

- Gmail address
- Gmail App Password
- Backend choice
- OpenAI API key if you choose OpenAI
- Writing everything into your local gitignored `.env.local`

For most people, choose OpenAI when prompted. It is the easiest option because it does not require a local Ollama installation.

If you prefer to create the file manually, copy the example first:

```bash
cp .env.local.example .env.local
```

Or on PowerShell:

```powershell
Copy-Item .env.local.example .env.local
```

Optional: create a personal allowlist file for senders or domains you always want to keep:

```bash
cp allowlist.example.txt allowlist.txt
```

Required values for `.env.local`:

- `EMAIL_CLASSIFIER_BACKEND` set to `openai` or `ollama`
- `GMAIL_EMAIL`
- `GMAIL_APP_PASSWORD`
- `OPENAI_API_KEY` when using `--backend openai`

Useful optional values:

- `OPENAI_MODEL`
- `OLLAMA_URL`
- `OLLAMA_MODEL`

If required settings are missing and you run the CLI in an interactive terminal, it will prompt you and offer to save the values into `.env.local`.

## Usage

Recommended first run after setup:

```bash
python main.py --folder "[Gmail]/All Mail" --gmail-query "category:promotions" --limit 50 --dry-run
```

OpenAI dry run with explicit model override:

```bash
python main.py --backend openai --openai-model gpt-4.1-mini --folder "[Gmail]/All Mail" --gmail-query "category:promotions" --limit 200 --dry-run
```

Generic dry run against Gmail Promotions via query:

```bash
python main.py --mode aggressive --folder "[Gmail]/All Mail" --gmail-query "category:promotions" --limit 200 --dry-run
```

Ollama backend:

```bash
python main.py --backend ollama --ollama-url http://127.0.0.1:11434 --ollama-model llama3:latest --mode aggressive --folder "[Gmail]/All Mail" --gmail-query "category:promotions" --limit 500
```

OpenAI backend:

```bash
python main.py --backend openai --openai-model gpt-4.1-mini --folder "Promotions" --limit 200
```

Ollama OpenAI-compatible endpoint:

```bash
python main.py --backend ollama --ollama-url http://127.0.0.1:11434/v1 --ollama-model llama3.1:8b --limit 200
```

If Promotions is an Inbox tab rather than an IMAP folder:

```bash
python main.py --folder "[Gmail]/All Mail" --gmail-query "category:promotions" --limit 200 --dry-run
```

Reprocess messages that were already marked as handled:

```bash
python main.py --mode aggressive --folder "[Gmail]/All Mail" --limit 1500 --no-skip-processed
```

Remove AI labels in bulk:

```bash
python main.py --cleanup-ai-labels --folder "[Gmail]/All Mail" --limit 150000
```

Remove AI labels plus the processed label:

```bash
python main.py --cleanup-ai-labels --cleanup-include-processed --folder "[Gmail]/All Mail" --limit 150000
```

Use a custom allowlist path:

```bash
python main.py --allowlist-path "./allowlist.txt" --mode aggressive --limit 200
```

## Behavior Notes

- `--mode aggressive` is the default.
- `--setup` launches the guided first-run wizard and writes `.env.local`.
- `EMAIL_CLASSIFIER_BACKEND` lets `.env.local` choose the default backend, so most users do not need to remember `--backend`.
- In aggressive mode, deterministic rules run before the LLM.
- Confidence below `0.85` is forced to `REVIEW`.
- Promotions scans can be delete-first via `--folder-default-policy promotions-delete`.
- Human-looking senders with no unsubscribe header are protected by a keep safeguard.
- `--dry-run` skips label creation and only writes the audit CSV.
- Newest messages are processed first.
- `--skip-processed` is enabled by default.
- `--processed-label` lets you override `AI/PROCESSED`.

## Allowlist Format

`allowlist.txt` is local-only and gitignored. Example:

```text
# Exact email
friend@example.com

# Whole domain
example.org
```

## Tests

Run the fixture-based policy test:

```bash
python -m unittest tests/test_policy_fixtures.py
```

## Publishing Checklist

Before your first public commit:

1. Fill in `.env.local` with your own credentials.
2. Put any personal sender rules in `allowlist.txt`.
3. Check `git status` and confirm only shareable files are staged.
4. Do not force-add ignored files.

If you want to publish this repo to GitHub after reviewing the file list:

```bash
git add .
git commit -m "Prepare repository for sharing"
git remote add origin <your-repo-url>
git push -u origin main
```

## Prompt for an Agent

Use this prompt if you want a coding agent to set up this repo for a non-developer on their machine, step by step:

```text
Set up this repository for me locally in the simplest possible way.

Assume I am a designer, not a developer.

Requirements:
- Install what is needed.
- Use the project's guided setup flow if it exists.
- If credentials or API keys are needed, ask me for them one at a time in plain English.
- Store secrets only in local gitignored files, not in tracked files.
- Prefer the easiest backend for a non-technical user.
- After setup, run a safe dry run command so we can confirm the tool works without changing live data.
- Explain each step briefly and tell me exactly what you need from me when input is required.
- If a command fails, diagnose it and continue instead of stopping at the first error.
```

Use this prompt if you want a coding agent to make another repository safe to publish in the same way:

```text
Take this repository and prepare it for public sharing.

Requirements:
- Find any local secrets, API keys, personal config, audit exports, caches, or browser profile data that should not be committed.
- Move runtime credentials into a gitignored .env.local file.
- Add or improve an interactive setup flow so a non-developer can be guided through first-run setup.
- Add a tracked .env.local.example file with placeholders.
- Add or update .gitignore so local-only files stay out of version control.
- If the app already reads environment variables, preserve that behavior. If it does not, add minimal support for loading .env.local without breaking existing CLI usage.
- Add tracked example files for any personal local config, such as allowlists.
- Review and rewrite the README so a new user can install, configure, run, test, and publish the repo safely.
- End the README with a short section that explains which files are local-only and which files are safe to commit.
- Initialize git if the folder is not already a repository.
- Validate the result with the narrowest relevant tests or checks and summarize any manual follow-up.

Constraints:
- Make the smallest practical changes.
- Do not commit real secrets.
- Do not delete user data unless necessary; prefer ignoring it.
- Preserve the current project behavior unless a change is required to support .env.local.
```

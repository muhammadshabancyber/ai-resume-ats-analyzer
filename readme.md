# ai-resume-ats-analyzer
# 📄 AI Resume ATS Checker

A simple web app that gives your resume an **estimated ATS score** and clear, actionable tips to improve it. Built with [Streamlit](https://streamlit.io) and the Google Gemini API.

> **Important:** The score is an AI-based *estimate*. Real Applicant Tracking Systems differ from each other, and no tool can give an official score.

## Features

- Upload a **PDF** or **DOCX** resume (up to 5 MB).
- Estimated overall ATS score plus scores for **Structure, Formatting, Skills, Experience, Education and Keywords**.
- A short list of the most important fixes, plus example bullet-point rewrites.
- **Optional job description**: match score, matched and missing keywords, and tailoring tips.
- Download the report as a Markdown file.
- Friendly messages for invalid files, scanned (image-only) PDFs, password-protected PDFs, a missing API key and API problems.

## Files

| File | What it is |
| --- | --- |
| `app.py` | The whole application |
| `requirements.txt` | Python packages the app needs |
| `README.md` | This guide |

## Get a Gemini API key

1. Go to <https://aistudio.google.com/apikey> and sign in with a Google account.
2. Click **Create API key** and copy it.
3. Treat it like a password. **Never** paste it into `app.py`, the README, or any file you upload to GitHub.

## Run it on your own computer (optional)

You need Python 3.10 or newer.

```bash
# 1. Open a terminal in the folder that contains the three files
python -m venv venv

# 2. Activate the virtual environment
#    macOS / Linux:
source venv/bin/activate
#    Windows (PowerShell):
venv\Scripts\Activate.ps1

# 3. Install the packages
pip install -r requirements.txt

# 4. Set your API key for this terminal session
#    macOS / Linux:
export GEMINI_API_KEY="paste-your-key-here"
#    Windows (PowerShell):
$env:GEMINI_API_KEY="paste-your-key-here"

# 5. Start the app
streamlit run app.py
```

Your browser opens the app automatically (usually at <http://localhost:8501>).

## Deploy on Streamlit Community Cloud (free)

1. Create a GitHub repository and upload the three files (`app.py`, `requirements.txt`, `README.md`) to its main page, not inside a folder.
2. Go to <https://share.streamlit.io> and sign in with GitHub.
3. Click **Create app** → **Yup, I have an app**.
4. Choose your repository, the `main` branch, and set the main file path to `app.py`.
5. Click **Advanced settings**, choose **Python 3.12**, and paste this into **Secrets**:

   ```toml
   GEMINI_API_KEY = "paste-your-key-here"
   ```

6. Click **Save**, then deploy. The first start takes a few minutes.

To change a secret later, open your app's settings on Streamlit Community Cloud and edit **Secrets**, then reboot the app.

## Settings (all optional, set in Secrets)

| Secret | Purpose |
| --- | --- |
| `GEMINI_API_KEY` | **Required.** Your Gemini API key. |
| `GEMINI_MODEL` | Use a different Gemini model name. Default: `gemini-3.7-flash`. Change this if Google renames or retires a model. |
| `APP_PASSWORD` | If set, visitors must enter this password before using the app. Recommended if your app link is public, because every analysis uses **your** API quota. |

Example:

```toml
GEMINI_API_KEY = "paste-your-key-here"
GEMINI_MODEL = "gemini-3.7-flash"
APP_PASSWORD = "choose-a-strong-password"
```

## Limits built into the app

- Files larger than 5 MB are rejected. Only the first 10 PDF pages are analyzed.
- Resume text is cut at 30,000 characters and the job description at 10,000 characters.
- There is a 15-second wait between analyses in one browser session.
- Temporary Gemini errors (busy server, rate limit) are retried automatically up to 2 times.

## Privacy and security

- Your resume text is sent to Google's Gemini API for analysis. Do not upload anything you are not comfortable sharing. Check Google's current terms for how API data is handled, especially on the free tier.
- The app does not save uploaded files or resume text, and it does not print them to logs.
- The API key is read only from Streamlit Secrets or an environment variable.
- Text inside resumes and job descriptions is treated as data, and the AI is told not to follow instructions written inside it.

## Troubleshooting

| Problem | What to do |
| --- | --- |
| "The Gemini API key is missing" | Add `GEMINI_API_KEY` in Secrets, save, and reboot the app. |
| "Gemini rejected the API key" | Create a new key, paste it with no spaces or extra quotes, and make sure the Gemini API is enabled for it. |
| "The Gemini model was not found" | Set `GEMINI_MODEL` in Secrets to a current Gemini Flash model name from Google's model list. |
| "Usage limit was reached" | Wait a few minutes. Free keys have low limits. |
| "Almost no readable text" | Your file is probably a scanned image. Export a text-based PDF or DOCX from Word or Google Docs. |
| App fails to install packages | In **Advanced settings** pick Python 3.12, then reboot the app. |

## License

Use and modify this project however you like.

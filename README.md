# Bulk Certificate Mailer

A simple Python desktop application that:

- Uses a Canva-exported certificate PNG as the background.
- Reads recipients from Excel or CSV.
- Adds each person's name automatically.
- Optionally adds a Certificate ID.
- Generates a PDF for every participant.
- Sends each participant an individual email with their certificate attached.
- Uses Gmail SMTP with a Gmail App Password.

## 1. Install

```bash
python -m venv .venv
```

Windows:
```bash
.venv\Scripts\activate
```

Linux/macOS:
```bash
source .venv/bin/activate
```

```bash
pip install -r requirements.txt
```

## 2. Prepare your participant Excel

Required columns:

| Name | Email |
|---|---|
| Rahul Kumar | rahul@example.com |
| Priya Sharma | priya@example.com |

Optional:
- Certificate_ID

## 3. Canva template

Export the certificate from Canva as PNG at the highest available quality.

In the app:
1. Choose Certificate PNG.
2. Choose your Excel/CSV.
3. Click Preview First Certificate.
4. Adjust Name X / Name Y / Font size if needed.
5. Generate All Certificates.

The initial coordinates are calibrated for the supplied sample certificate.

## 4. Gmail

For Gmail, enable 2-Step Verification and create a Google App Password.
Use the 16-character App Password in the application, not your normal Gmail password.

## 5. Send

Enter:
- Gmail address
- Gmail App Password
- Subject
- Email body

Use `{name}` in the email body to insert the participant's name.

Then click Generate + Send All.

Generated certificates are stored in:

`generated_certificates/`

## Important

The application sends one separate email per recipient. It does not put all recipients in CC/BCC, so participants do not see one another's email addresses.

For large distributions, check your organization's/email provider's sending limits before sending.

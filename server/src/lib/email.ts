import nodemailer from 'nodemailer'

const smtpPort = parseInt(process.env.SMTP_PORT || '587')
const transporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST,
  port: smtpPort,
  secure: smtpPort === 465,  // true for 465 (SSL), false for 587 (STARTTLS)
  auth: {
    user: process.env.SMTP_USER,
    pass: process.env.SMTP_PASS,
  },
  connectionTimeout: 30000,
  greetingTimeout: 30000,
  socketTimeout: 30000,
})

export async function sendReviewInvite(
  to: string,
  reviewLink: string,
  projectName: string,
): Promise<void> {
  await transporter.sendMail({
    from: process.env.SMTP_FROM,
    to,
    subject: `ESD Review Invitation – ${projectName}`,
    html: `
      <div style="font-family: 'Open Sans', Arial, sans-serif; color: #2C2C2C; max-width: 600px;">
        <h2 style="font-family: Montserrat, Arial, sans-serif; color: #6B7A3B;">
          ESD Review Invitation
        </h2>
        <p>You have been invited to review the ESD credits for <strong>${projectName}</strong>.</p>
        <p>
          <a
            href="${reviewLink}"
            style="display:inline-block; background:#6B7A3B; color:#fff; padding:10px 22px;
                   text-decoration:none; font-family:Montserrat,Arial,sans-serif; font-weight:500;
                   border-radius:2px;"
          >
            Open My Review
          </a>
        </p>
        <p style="color:#8C8C8C; font-size:13px;">
          This link is unique to you — please do not share it.
        </p>
      </div>
    `,
  })
}

export async function sendMagicLink(to: string, link: string): Promise<void> {
  await transporter.sendMail({
    from: process.env.SMTP_FROM,
    to,
    subject: 'Your GIW login link',
    html: `
      <div style="font-family: 'Open Sans', Arial, sans-serif; color: #2C2C2C; max-width: 600px;">
        <h2 style="font-family: Montserrat, Arial, sans-serif; color: #6B7A3B;">
          GIW Environmental Solutions
        </h2>
        <p>Use the button below to sign in. This link expires in 15 minutes.</p>
        <p>
          <a
            href="${link}"
            style="display:inline-block; background:#6B7A3B; color:#fff; padding:10px 22px;
                   text-decoration:none; font-family:Montserrat,Arial,sans-serif; font-weight:500;
                   border-radius:2px;"
          >
            Sign In
          </a>
        </p>
        <p style="color:#8C8C8C; font-size:13px;">
          If you did not request this link, you can ignore this email.
        </p>
      </div>
    `,
  })
}

export async function sendReviewSubmission(opts: {
  toEmail: string
  discipline: string
  projectName: string
  projectId: string
  submittedAt: Date
}): Promise<void> {
  const base = process.env.BASE_URL || 'http://localhost:5173'
  const adminLink = `${base}/admin/projects/${opts.projectId}`
  const dateStr = opts.submittedAt.toLocaleString('en-AU', {
    timeZone: 'Australia/Melbourne',
    dateStyle: 'medium',
    timeStyle: 'short',
  })

  await transporter.sendMail({
    from: process.env.SMTP_FROM,
    to: 'info@giw.com.au',
    subject: `${opts.discipline} review submitted – ${opts.projectName}`,
    html: `
      <div style="font-family: 'Open Sans', Arial, sans-serif; color: #2C2C2C; max-width: 600px;">
        <h2 style="font-family: Montserrat, Arial, sans-serif; color: #6B7A3B;">
          Review Submitted
        </h2>
        <p>
          ${opts.toEmail} (${opts.discipline}) submitted their ESD review for
          <strong>${opts.projectName}</strong> on ${dateStr}.
        </p>
        <p>
          <a href="${adminLink}"
             style="display:inline-block; background:#6B7A3B; color:#fff; padding:10px 22px;
                    text-decoration:none; font-family:Montserrat,Arial,sans-serif; font-weight:500;
                    border-radius:2px;">
            Open Project
          </a>
        </p>
      </div>
    `,
  })
}

export async function sendReviewInviteByEmail(
  to: string,
  inviteLink: string,
  projectName: string,
  discipline: string,
  recipientName: string | null,
  projectAddress: string | null,
): Promise<void> {
  const BASE_URL = process.env.BASE_URL || 'http://localhost:5173'
  const logoUrl = `${BASE_URL}/GIW%20logo.png`
  const greeting = recipientName ? `Hi ${recipientName},` : 'Hi,'
  const addressLine = projectAddress ?? projectName

  await transporter.sendMail({
    from: process.env.SMTP_FROM,
    to,
    subject: `ESD Review Invitation – ${addressLine}`,
    html: `
      <!DOCTYPE html>
      <html>
      <head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
      <body style="margin:0;padding:0;background:#F7F5F0;">
        <table width="100%" cellpadding="0" cellspacing="0" style="background:#F7F5F0;padding:32px 0;">
          <tr><td align="center">
            <table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#ffffff;border:1px solid #D8D5CE;border-radius:4px;overflow:hidden;">
              <!-- Header -->
              <tr>
                <td style="background:#2C2C2C;padding:20px 32px;">
                  <img src="${logoUrl}" alt="GIW Environmental Solutions" height="40" style="display:block;">
                </td>
              </tr>
              <!-- Body -->
              <tr>
                <td style="padding:40px 32px;font-family:'Open Sans',Arial,sans-serif;color:#2C2C2C;font-size:15px;line-height:1.6;">
                  <p style="margin:0 0 16px 0;">${greeting}</p>
                  <p style="margin:0 0 24px 0;">You have been invited to provide comments on the ESD review for the proposed development at <strong>${addressLine}</strong>.</p>
                  <p style="margin:0 0 24px 0;">Click on the following link to log into the reviewer portal.</p>
                  <p style="margin:0 0 32px 0;">
                    <a href="${inviteLink}"
                       style="display:inline-block;background:#6B7A3B;color:#ffffff;padding:12px 28px;
                              text-decoration:none;font-family:Montserrat,Arial,sans-serif;font-weight:500;
                              font-size:14px;border-radius:2px;letter-spacing:0.3px;">
                      Access Reviewer Portal
                    </a>
                  </p>
                </td>
              </tr>
              <!-- Footer -->
              <tr>
                <td style="background:#E8EDD8;padding:16px 32px;font-family:'Open Sans',Arial,sans-serif;font-size:12px;color:#8C8C8C;">
                  GIW Environmental Solutions &nbsp;|&nbsp; <a href="https://giw.com.au" style="color:#6B7A3B;text-decoration:none;">giw.com.au</a>
                </td>
              </tr>
            </table>
          </td></tr>
        </table>
      </body>
      </html>
    `,
  })
}

export async function sendSubmissionAlert(opts: {
  submitterEmail: string
  submitterDiscipline: string
  projectName: string
  projectId: string
  submittedAt: Date
  reviewerEmails: string[]
  notifyEmail: string | null
}): Promise<void> {
  const emails = [
    ...opts.reviewerEmails,
    ...(opts.notifyEmail ? [opts.notifyEmail] : []),
  ].filter(Boolean)
  if (emails.length === 0) return

  const base = process.env.BASE_URL || 'http://localhost:5173'
  const adminLink = `${base}/admin/projects/${opts.projectId}`
  const dateStr = opts.submittedAt.toLocaleString('en-AU', {
    timeZone: 'Australia/Melbourne',
    dateStyle: 'medium',
    timeStyle: 'short',
  })

  await transporter.sendMail({
    from: process.env.SMTP_FROM,
    to: emails.join(', '),
    subject: `${opts.submitterDiscipline} review submitted – ${opts.projectName}`,
    html: `
      <div style="font-family: 'Open Sans', Arial, sans-serif; color: #2C2C2C; max-width: 600px;">
        <h2 style="font-family: Montserrat, Arial, sans-serif; color: #6B7A3B;">
          Review Submitted
        </h2>
        <p>
          ${opts.submitterEmail} (${opts.submitterDiscipline}) has submitted their ESD review for
          <strong>${opts.projectName}</strong> on ${dateStr}.
        </p>
        <p>
          <a href="${adminLink}"
             style="display:inline-block; background:#00602B; color:#fff; padding:10px 22px;
                    text-decoration:none; font-family:Montserrat,Arial,sans-serif; font-weight:500;
                    border-radius:2px;">
            Open Project
          </a>
        </p>
      </div>
    `,
  })
}

export async function sendCompletionEmail(opts: {
  projectName: string
  reviewerEmails: string[]
  excelBuffer: Buffer
  fileName: string
}): Promise<void> {
  await transporter.sendMail({
    from: process.env.SMTP_FROM,
    to: opts.reviewerEmails.join(', '),
    subject: `ESD Review Complete – ${opts.projectName}`,
    html: `
      <div style="font-family: 'Open Sans', Arial, sans-serif; color: #2C2C2C; max-width: 600px;">
        <h2 style="font-family: Montserrat, Arial, sans-serif; color: #6B7A3B;">
          ESD Review Complete
        </h2>
        <p>
          All reviewers have now submitted their comments for <strong>${opts.projectName}</strong>.
          Please find the complete ESD Review Matrix attached.
        </p>
      </div>
    `,
    attachments: [
      {
        filename: opts.fileName,
        content: opts.excelBuffer,
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
    ],
  })
}

export async function sendSubmissionNotification(
  projectName: string,
  reviewerEmail: string,
  reviewerDiscipline: string,
): Promise<void> {
  const to = process.env.GIW_NOTIFY_EMAIL
  if (!to) return

  await transporter.sendMail({
    from: process.env.SMTP_FROM,
    to,
    subject: `Review submitted – ${projectName}`,
    html: `
      <div style="font-family: 'Open Sans', Arial, sans-serif; color: #2C2C2C;">
        <h2 style="font-family: Montserrat, Arial, sans-serif;">Review Submitted</h2>
        <p>
          <strong>${reviewerEmail}</strong> (${reviewerDiscipline}) has submitted their
          review for <strong>${projectName}</strong>.
        </p>
      </div>
    `,
  })
}

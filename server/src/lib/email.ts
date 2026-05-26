import { Resend } from 'resend'

const resend = new Resend(process.env.RESEND_API_KEY)
const FROM = process.env.SMTP_FROM || 'GIW Environmental Solutions <onboarding@resend.dev>'

async function send(to: string | string[], subject: string, html: string, attachments?: Array<{ filename: string; content: Buffer; contentType: string }>) {
  await resend.emails.send({
    from: FROM,
    to: Array.isArray(to) ? to : [to],
    subject,
    html,
    attachments: attachments?.map(a => ({
      filename: a.filename,
      content: a.content,
    })),
  })
}

/* ── Shared branded wrapper ── */
function emailTemplate(bodyContent: string): string {
  const BASE_URL = process.env.BASE_URL || 'http://localhost:5173'
  const logoUrl = `${BASE_URL}/GIW%20logo.png`
  return `<!DOCTYPE html>
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
          <td style="padding:40px 32px;font-family:'Open Sans',Arial,sans-serif;color:#2C2C2C;font-size:15px;line-height:1.7;">
            ${bodyContent}
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
</html>`
}

function btn(href: string, label: string, color = '#6B7A3B'): string {
  return `<p style="margin:28px 0 0 0;">
    <a href="${href}" style="display:inline-block;background:${color};color:#ffffff;padding:12px 28px;
       text-decoration:none;font-family:Montserrat,Arial,sans-serif;font-weight:500;
       font-size:14px;border-radius:2px;letter-spacing:0.3px;">${label}</a>
  </p>`
}

/* ── Email functions ── */

export async function sendReviewInvite(
  to: string,
  reviewLink: string,
  projectName: string,
): Promise<void> {
  await send(to, `ESD Review Invitation – ${projectName}`, emailTemplate(`
    <p style="margin:0 0 16px 0;">Hi,</p>
    <p style="margin:0 0 16px 0;">You have been invited to review the ESD credits for <strong>${projectName}</strong>.</p>
    ${btn(reviewLink, 'Open My Review')}
    <p style="margin:24px 0 0 0;color:#8C8C8C;font-size:13px;">This link is unique to you — please do not share it.</p>
  `))
}

export async function sendMagicLink(to: string, link: string): Promise<void> {
  await send(to, 'Your GIW login link', emailTemplate(`
    <p style="margin:0 0 16px 0;">Hi,</p>
    <p style="margin:0 0 16px 0;">Use the button below to sign in to the GIW ESD Review Portal. This link expires in 15 minutes.</p>
    ${btn(link, 'Sign In')}
    <p style="margin:24px 0 0 0;color:#8C8C8C;font-size:13px;">If you did not request this link, you can ignore this email.</p>
  `))
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
  await send('info@giw.com.au', `${opts.discipline} review submitted – ${opts.projectName}`, emailTemplate(`
    <p style="margin:0 0 16px 0;">Hi,</p>
    <p style="margin:0 0 16px 0;">
      ${opts.toEmail} (${opts.discipline}) submitted their ESD review for
      <strong>${opts.projectName}</strong> on ${dateStr}.
    </p>
    ${btn(adminLink, 'Open Project', '#00602B')}
  `))
}

export async function sendReviewInviteByEmail(
  to: string,
  inviteLink: string,
  projectName: string,
  discipline: string,
  recipientName: string | null,
  projectAddress: string | null,
): Promise<void> {
  const greeting = recipientName ? `Hi ${recipientName},` : 'Hi,'
  const addressLine = projectAddress ?? projectName

  await send(to, `ESD Review Invitation – ${addressLine}`, emailTemplate(`
    <p style="margin:0 0 16px 0;">${greeting}</p>
    <p style="margin:0 0 16px 0;">
      GIW Environmental Solutions have undertaken an ESD Review for the proposed development at
      <strong>${addressLine}</strong>. You have been invited to provide comments on the ESD review
      via the link below.
    </p>
    <p style="margin:0;">Click on the following link to log into the reviewer portal.</p>
    ${btn(inviteLink, 'Access Reviewer Portal')}
  `))
}

export async function sendSubmissionAlert(opts: {
  submitterEmail: string
  submitterDiscipline: string
  projectName: string
  projectId: string
  reviewLinkToken: string
  submittedAt: Date
  reviewerEmails: string[]
  notifyEmail: string | null
}): Promise<void> {
  const reviewerEmails = opts.reviewerEmails.filter(Boolean)
  if (reviewerEmails.length === 0 && !opts.notifyEmail) return

  const base = process.env.BASE_URL || 'http://localhost:5173'
  const adminLink = `${base}/admin/projects/${opts.projectId}`
  const reviewLink = `${base}/review/${opts.reviewLinkToken}`
  const dateStr = opts.submittedAt.toLocaleString('en-AU', {
    timeZone: 'Australia/Melbourne',
    dateStyle: 'medium',
    timeStyle: 'short',
  })

  const subject = `${opts.submitterDiscipline} review submitted – ${opts.projectName}`

  const reviewerHtml = emailTemplate(`
    <p style="margin:0 0 16px 0;">Hi,</p>
    <p style="margin:0 0 16px 0;">
      ${opts.submitterEmail} (${opts.submitterDiscipline}) has submitted their ESD review for
      <strong>${opts.projectName}</strong> on ${dateStr}.
    </p>
    ${btn(reviewLink, 'Open Review Portal')}
  `)

  const adminHtml = emailTemplate(`
    <p style="margin:0 0 16px 0;">Hi,</p>
    <p style="margin:0 0 16px 0;">
      ${opts.submitterEmail} (${opts.submitterDiscipline}) has submitted their ESD review for
      <strong>${opts.projectName}</strong> on ${dateStr}.
    </p>
    ${btn(adminLink, 'Open Project', '#00602B')}
  `)

  const sends: Promise<void>[] = []
  if (reviewerEmails.length > 0) sends.push(send(reviewerEmails, subject, reviewerHtml))
  if (opts.notifyEmail) sends.push(send(opts.notifyEmail, subject, adminHtml))
  await Promise.all(sends)
}

export async function sendCompletionEmail(opts: {
  projectName: string
  reviewerEmails: string[]
  excelBuffer: Buffer
  fileName: string
}): Promise<void> {
  await send(
    opts.reviewerEmails,
    `ESD Review Complete – ${opts.projectName}`,
    emailTemplate(`
      <p style="margin:0 0 16px 0;">Hi,</p>
      <p style="margin:0 0 16px 0;">
        All reviewers have now submitted their comments for <strong>${opts.projectName}</strong>.
        Please find the complete ESD Review Matrix attached.
      </p>
    `),
    [{ filename: opts.fileName, content: opts.excelBuffer, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }],
  )
}

export async function sendSubmissionNotification(
  projectName: string,
  reviewerEmail: string,
  reviewerDiscipline: string,
): Promise<void> {
  const to = process.env.GIW_NOTIFY_EMAIL
  if (!to) return

  await send(to, `Review submitted – ${projectName}`, emailTemplate(`
    <p style="margin:0 0 16px 0;">Hi,</p>
    <p style="margin:0;">
      <strong>${reviewerEmail}</strong> (${reviewerDiscipline}) has submitted their
      review for <strong>${projectName}</strong>.
    </p>
  `))
}

import { graphPost } from "./graphService.js";

// STEP 4: Send via Graph sendMail using the mailbox in SENDER_EMAIL.
// App registration needs Mail.Send (application) consented for that mailbox,
// scoped via an Exchange Application Access Policy for production.
export async function sendAlertEmail(
  recipient: string,
  listIdOrTitle: string,
  changeType: "New items" | "Modified" | "Deleted",
  itemId: number,
  when: string
): Promise<void> {
  const sender = process.env.SENDER_EMAIL!;
  const subject = `Alert: ${listIdOrTitle} changed`;

  // Inline HTML — keep readable on a projector. No template engine.
  const html = `
    <p>Hi,</p>
    <p>A change was detected on the list you're watching:</p>
    <table border="1" cellpadding="6" cellspacing="0">
      <tr><th>Item</th><th>Change type</th><th>Time</th></tr>
      <tr>
        <td>#${itemId}</td>
        <td>${changeType}</td>
        <td>${new Date(when).toLocaleString()}</td>
      </tr>
    </table>
    <p style="color:#888;font-size:12px">Sent by the TechEd 2026 SP Webhook demo.</p>
  `;

  await graphPost(`/users/${sender}/sendMail`, {
    message: {
      subject,
      body: { contentType: "HTML", content: html },
      toRecipients: [{ emailAddress: { address: recipient } }],
    },
    saveToSentItems: false,
  });
}

import { Resend } from 'resend';
import admin from 'firebase-admin';
import { initializeFirebase } from './firebase';

const resend = new Resend(process.env.RESEND_API_KEY);

const FROM_ADDRESS = 'GridPath <support@gridpath.dev>';
const APP_URL = 'https://gridpath.dev';

const WELCOME_EMAIL_HTML = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin: 0; padding: 0; font-size: 14px; line-height: 24px; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Fira Sans', 'Droid Sans', 'Helvetica Neue', sans-serif; color: #333;">
  <div style="padding: 20px; max-width: 480px; margin: 0 auto;">
    <p style="margin: 0; padding-top: 0.5em; padding-bottom: 0.5em;">Hey,</p>

    <p style="margin: 0; padding-top: 0.5em; padding-bottom: 0.5em;">Welcome to GridPath — Cursor for Excel.</p>

    <p style="margin: 0; padding-top: 0.5em; padding-bottom: 0.5em;">Open GridPath, then open a spreadsheet in Excel and just describe what you want — build a model, clean data, write formulas, lay out a sheet. GridPath does it.</p>

    <p style="margin: 0; padding-top: 0.5em; padding-bottom: 0.5em;">A few things to try:</p>

    <p style="margin: 0; padding-top: 0.5em; padding-bottom: 0.5em;">→ "Build a 3-statement DCF for [company] with 5-year projections"<br>
    → "Clean up this sheet — fix the headers, standardize the date column, total each section"<br>
    → "Pivot this list of transactions by category and month"</p>

    <p style="margin: 0; padding-top: 0.5em; padding-bottom: 0.5em;">Questions? Just reply to this email.</p>
  </div>
</body>
</html>
`;

export async function sendWelcomeEmail(userEmail: string): Promise<boolean> {
  if (!process.env.RESEND_API_KEY) {
    console.warn('RESEND_API_KEY not configured, skipping welcome email');
    return false;
  }

  try {
    const { error } = await resend.emails.send({
      from: FROM_ADDRESS,
      to: userEmail,
      subject: 'Welcome to GridPath',
      html: WELCOME_EMAIL_HTML,
    });

    if (error) {
      console.error('Failed to send welcome email:', error);
      return false;
    }

    console.log(`Welcome email sent to ${userEmail}`);
    return true;
  } catch (error) {
    console.error('Error sending welcome email:', error);
    return false;
  }
}

// Verification email template
const getVerificationEmailHtml = (verificationLink: string) => `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin: 0; padding: 0; font-size: 14px; line-height: 24px; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Fira Sans', 'Droid Sans', 'Helvetica Neue', sans-serif; color: #333; background-color: #f9fafb;">
  <div style="max-width: 480px; margin: 0 auto; padding: 40px 20px;">
    <div style="background: white; border-radius: 12px; padding: 32px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
      <h1 style="margin: 0 0 16px 0; font-size: 20px; font-weight: 600; text-align: center;">Verify your email</h1>

      <p style="margin: 0 0 24px 0; color: #666; text-align: center;">
        Click the button below to verify your email and start using GridPath.
      </p>

      <div style="text-align: center; margin-bottom: 24px;">
        <a href="${verificationLink}" style="display: inline-block; background-color: #2563eb; color: white; text-decoration: none; padding: 12px 32px; border-radius: 8px; font-weight: 500;">
          Verify Email
        </a>
      </div>

      <p style="margin: 0; font-size: 12px; color: #999; text-align: center;">
        This link expires in 1 hour. If you didn't create an account, you can ignore this email.
      </p>
    </div>

    <p style="margin: 24px 0 0 0; font-size: 12px; color: #999; text-align: center;">
      © ${new Date().getFullYear()} GridPath
    </p>
  </div>
</body>
</html>
`;

export async function sendVerificationEmail(userEmail: string): Promise<{ success: boolean; error?: string }> {
  if (!process.env.RESEND_API_KEY) {
    console.warn('RESEND_API_KEY not configured, skipping verification email');
    return { success: false, error: 'Email service not configured' };
  }

  try {
    initializeFirebase();

    // Firebase generates a signed action URL — clicking it marks the email
    // verified. Without a custom Firebase auth domain, the link uses the
    // default `<project>.firebaseapp.com` host, which works as-is.
    const verificationLink = await admin.auth().generateEmailVerificationLink(userEmail, {
      url: APP_URL,
    });

    const { error } = await resend.emails.send({
      from: FROM_ADDRESS,
      to: userEmail,
      subject: 'Verify your email — GridPath',
      html: getVerificationEmailHtml(verificationLink),
    });

    if (error) {
      console.error('Failed to send verification email:', error);
      return { success: false, error: 'Failed to send email' };
    }

    console.log(`Verification email sent to ${userEmail}`);
    return { success: true };
  } catch (error: any) {
    console.error('Error sending verification email:', error);
    return { success: false, error: error.message || 'Failed to send verification email' };
  }
}

// Password reset email template
const getPasswordResetEmailHtml = (resetLink: string) => `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin: 0; padding: 0; font-size: 14px; line-height: 24px; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Fira Sans', 'Droid Sans', 'Helvetica Neue', sans-serif; color: #333; background-color: #f9fafb;">
  <div style="max-width: 480px; margin: 0 auto; padding: 40px 20px;">
    <div style="background: white; border-radius: 12px; padding: 32px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
      <h1 style="margin: 0 0 16px 0; font-size: 20px; font-weight: 600; text-align: center;">Reset your password</h1>

      <p style="margin: 0 0 24px 0; color: #666; text-align: center;">
        Click the button below to reset your GridPath password.
      </p>

      <div style="text-align: center; margin-bottom: 24px;">
        <a href="${resetLink}" style="display: inline-block; background-color: #2563eb; color: white; text-decoration: none; padding: 12px 32px; border-radius: 8px; font-weight: 500;">
          Reset Password
        </a>
      </div>

      <p style="margin: 0; font-size: 12px; color: #999; text-align: center;">
        This link expires in 1 hour. If you didn't request a password reset, you can ignore this email.
      </p>
    </div>

    <p style="margin: 24px 0 0 0; font-size: 12px; color: #999; text-align: center;">
      © ${new Date().getFullYear()} GridPath
    </p>
  </div>
</body>
</html>
`;

export async function sendPasswordResetEmail(userEmail: string): Promise<{ success: boolean; error?: string }> {
  if (!process.env.RESEND_API_KEY) {
    console.warn('RESEND_API_KEY not configured, skipping password reset email');
    return { success: false, error: 'Email service not configured' };
  }

  try {
    initializeFirebase();

    const resetLink = await admin.auth().generatePasswordResetLink(userEmail, {
      url: APP_URL,
    });

    const { error } = await resend.emails.send({
      from: FROM_ADDRESS,
      to: userEmail,
      subject: 'Reset your password — GridPath',
      html: getPasswordResetEmailHtml(resetLink),
    });

    if (error) {
      console.error('Failed to send password reset email:', error);
      return { success: false, error: 'Failed to send email' };
    }

    console.log(`Password reset email sent to ${userEmail}`);
    return { success: true };
  } catch (error: any) {
    console.error('Error sending password reset email:', error);
    return { success: false, error: error.message || 'Failed to send password reset email' };
  }
}

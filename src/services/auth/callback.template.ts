const callbackTemplate = `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>GridPath — Authentication Successful</title>
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
      display: flex;
      justify-content: center;
      align-items: center;
      height: 100vh;
      margin: 0;
      background: radial-gradient(ellipse at top, #1a1a1a 0%, #0a0a0a 70%);
      color: #e4e4e4;
    }
    .container {
      text-align: center;
      padding: 48px 56px;
      background: rgba(255, 255, 255, 0.04);
      border: 1px solid rgba(255, 255, 255, 0.08);
      border-radius: 16px;
      backdrop-filter: blur(10px);
      box-shadow: 0 8px 32px rgba(0, 0, 0, 0.4);
    }
    h1 {
      font-size: 1.75rem;
      font-weight: 600;
      margin: 0 0 12px 0;
      letter-spacing: -0.01em;
    }
    p {
      font-size: 0.95rem;
      color: #7c7c7c;
      margin: 0;
    }
    .checkmark {
      width: 56px;
      height: 56px;
      border-radius: 50%;
      background-color: #2563eb;
      display: inline-block;
      margin-bottom: 24px;
      position: relative;
    }
    .checkmark::after {
      content: '';
      position: absolute;
      left: 20px;
      top: 12px;
      width: 11px;
      height: 22px;
      border: solid white;
      border-width: 0 3px 3px 0;
      transform: rotate(45deg);
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="checkmark"></div>
    <h1>Signed in to GridPath</h1>
    <p>You can close this window and return to the app.</p>
  </div>
  <script>
    setTimeout(() => {
      window.close();
    }, 2000);
  </script>
</body>
</html>
`;

export default callbackTemplate;

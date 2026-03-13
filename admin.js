<!DOCTYPE html>
<html lang="pl">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>World Food – Panel admin</title>
  <link rel="stylesheet" href="style.css">
</head>
<body>
  <div id="admin-view">
    <header>
      <img class="brand-logo" alt="World Food – Oferta" src="https://firebasestorage.googleapis.com/v0/b/pdf-creator-f7a8b.firebasestorage.app/o/CREATOR%20BASIC%2Fmaspo%20logo.png?alt=media&token=8fe33ebe-04d2-42cb-acb0-10379dbd7e11">
      <div class="header-actions">
        <button class="btn-outline" id="admin-back-btn">Powrót</button>
      </div>
    </header>

    <main>
      <section class="admin-panel">
        <h2>Panel admin</h2>
        <div class="login-form">
          <label class="login-label" for="admin-user-email">Email użytkownika</label>
          <input id="admin-user-email" type="email" placeholder="nowy@firma.com">
          <label class="login-label" for="admin-user-password">Hasło użytkownika</label>
          <input id="admin-user-password" type="password" placeholder="Minimum 6 znaków">
          <button id="admin-create-user" class="btn-outline">Utwórz konto</button>
        </div>
        <div id="admin-create-msg" class="login-error"></div>
      </section>
    </main>
  </div>

  <script src="https://www.gstatic.com/firebasejs/9.22.2/firebase-app-compat.js"></script>
  <script src="https://www.gstatic.com/firebasejs/9.22.2/firebase-auth-compat.js"></script>
  <script src="admin.js"></script>
</body>
</html>

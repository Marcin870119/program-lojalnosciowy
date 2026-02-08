const firebaseConfig = {
  apiKey: 'AIzaSyDwzVRS5W2lklGMLcZJn-YPCK9OtBQZ7bI',
  authDomain: 'pdf-creator-f7a8b.firebaseapp.com',
  projectId: 'pdf-creator-f7a8b',
  appId: '1:606744201676:web:6f8c1b2c323fbaf6f3b569'
};

const emailInput = document.getElementById('admin-user-email');
const passwordInput = document.getElementById('admin-user-password');
const createBtn = document.getElementById('admin-create-user');
const msg = document.getElementById('admin-create-msg');
const backBtn = document.getElementById('admin-back-btn');

if(!localStorage.getItem('is_admin')){
  window.location.href = 'index.html';
}

if(backBtn){
  backBtn.addEventListener('click', () => {
    window.location.href = 'index.html';
  });
}

if(typeof firebase !== 'undefined'){
  if(!firebase.apps.length){
    firebase.initializeApp(firebaseConfig);
  }
  if(!firebase.apps.some(a => a.name === 'adminApp')){
    firebase.initializeApp(firebaseConfig, 'adminApp');
  }
  const adminAuth = firebase.app('adminApp').auth();

  createBtn.addEventListener('click', async () => {
    const email = emailInput.value.trim();
    const pass = passwordInput.value;
    msg.textContent = '';
    if(!email || !pass){
      msg.textContent = 'Podaj email i hasło.';
      return;
    }
    if(pass.length < 6){
      msg.textContent = 'Hasło min. 6 znaków.';
      return;
    }
    createBtn.disabled = true;
    try{
      await adminAuth.createUserWithEmailAndPassword(email, pass);
      await adminAuth.signOut();
      msg.textContent = 'Konto utworzone.';
      emailInput.value = '';
      passwordInput.value = '';
    }catch(e){
      msg.textContent = e.code || 'Nie udało się utworzyć konta.';
    }finally{
      createBtn.disabled = false;
    }
  });
}

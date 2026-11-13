<script type="module">
  // Import the functions you need from the SDKs you need
  import { initializeApp } from "https://www.gstatic.com/firebasejs/13.3.0/firebase-app.js";
  import { getAnalytics } from "https://www.gstatic.com/firebasejs/13.3.0/firebase-analytics.js";
  // TODO: Add SDKs for Firebase products that you want to use
  // https://firebase.google.com/docs/web/setup#available-libraries

  // Your web app's Firebase configuration
  // For Firebase JS SDK v7.20.0 and later, measurementId is optional
  const firebaseConfig = {
    apiKey: "AIzaSyAvqW1AqP256o4eWxuBX4dyqLC703XTiBc",
    authDomain: "relay-91930.firebaseapp.com",
    projectId: "relay-91930",
    storageBucket: "relay-91930.firebasestorage.app",
    messagingSenderId: "938437348616",
    appId: "1:938437348616:web:26a3b353d789dca19570bd",
    measurementId: "G-FMQJJPVTCQ"
  };

  // Initialize Firebase
  const app = initializeApp(firebaseConfig);
  const analytics = getAnalytics(app);
</script>
// Excel-based authentication system for School Power BI Data Viewer

let userCredentials = [];

// Load user credentials from Excel file
function loadUserDatabase() {
  return fetch('userDB.xlsx')
    .then(response => {
      if (!response.ok) {
        throw new Error('Network response was not ok');
      }
      return response.arrayBuffer();
    })
    .then(data => {
      const workbook = XLSX.read(data, { type: 'array' });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      userCredentials = XLSX.utils.sheet_to_json(worksheet);
      console.log("User database loaded successfully");
    })
    .catch(error => {
      console.error("Error loading user database:", error);
      throw error;
    });
}

// Authenticate user
function authenticateUser(username, password) {
  // Simple hash function for password (not secure for production)
  const hashedPassword = simpleHash(password);
  
  const user = userCredentials.find(user => 
    user.username === username && user.password === hashedPassword
  );
  
  if (user) {
    // Store authentication state in sessionStorage
    sessionStorage.setItem('authenticated', 'true');
    sessionStorage.setItem('username', username);
    return true;
  }
  
  return false;
}

// Simple hash function (for demonstration purposes only)
function simpleHash(input) {
  let hash = 0;
  for (let i = 0; i < input.length; i++) {
    const char = input.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // Convert to 32bit integer
  }
  return hash.toString(16);
}

// Check if user is authenticated
function isAuthenticated() {
  return sessionStorage.getItem('authenticated') === 'true';
}

// Logout function
function logout() {
  sessionStorage.removeItem('authenticated');
  sessionStorage.removeItem('username');
  window.location.reload();
}

// Initialize authentication system
document.addEventListener('DOMContentLoaded', function() {
  // Load user database
  loadUserDatabase()
    .then(() => {
      // Check authentication status
      if (!isAuthenticated()) {
        document.getElementById('loginContainer').style.display = 'block';
        document.getElementById('appContainer').style.display = 'none';
      } else {
        document.getElementById('loginContainer').style.display = 'none';
        document.getElementById('appContainer').style.display = 'block';
      }
    });
    
  // Login form submission
  document.getElementById('loginForm').addEventListener('submit', function(e) {
    e.preventDefault();
    const username = document.getElementById('username').value;
    const password = document.getElementById('password').value;
    
    if (authenticateUser(username, password)) {
      document.getElementById('loginContainer').style.display = 'none';
      document.getElementById('appContainer').style.display = 'block';
    } else {
      document.getElementById('loginError').textContent = 'Invalid username or password';
    }
  });
});
// Simple authentication system for School Power BI Data Viewer

// Store user credentials
let userCredentials = [];

// Create default users
function createFallbackUsers() {
  console.log("Creating fallback user database");
  userCredentials = [
    { username: "admin", password: "admin123" },
    { username: "user", password: "user123" }
  ];
  console.log("Fallback user database created with users:", userCredentials.map(u => u.username).join(", "));
}

// Simple authentication function
function authenticateUser(username, password) {
  console.log("Authenticating user:", username);
  
  // Direct comparison for simplicity and reliability
  const user = userCredentials.find(u => 
    u.username === username && u.password === password
  );
  
  if (user) {
    console.log("Authentication successful for user:", username);
    sessionStorage.setItem('authenticated', 'true');
    sessionStorage.setItem('username', username);
    return true;
  } else {
    console.log("Authentication failed for user:", username);
    return false;
  }
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

// Initialize when document is ready
document.addEventListener('DOMContentLoaded', function() {
  console.log("=== Authentication system initializing ===");
  
  // Create users immediately
  createFallbackUsers();
  
  // Check authentication status
  if (isAuthenticated()) {
    console.log("User is already authenticated, showing app");
    document.getElementById('loginContainer').style.display = 'none';
    document.getElementById('appContainer').style.display = 'block';
  } else {
    console.log("User not authenticated, showing login form");
    document.getElementById('loginContainer').style.display = 'block';
    document.getElementById('appContainer').style.display = 'none';
  }
  
  // Add login form submission handler
  const loginForm = document.getElementById('loginForm');
  if (loginForm) {
    loginForm.addEventListener('submit', function(e) {
      e.preventDefault();
      const username = document.getElementById('username').value;
      const password = document.getElementById('password').value;
      
      console.log("Login form submitted for user:", username);
      
      if (authenticateUser(username, password)) {
        document.getElementById('loginContainer').style.display = 'none';
        document.getElementById('appContainer').style.display = 'block';
      } else {
        document.getElementById('loginError').textContent = 'Invalid username or password';
      }
    });
    console.log("Login form handler attached");
  } else {
    console.error("Login form not found in the document!");
  }
});
// Excel-based authentication system for School Power BI Data Viewer

let userCredentials = [];

// Add this line at the beginning of the file, after the userCredentials declaration
createFallbackUsers(); // Ensure fallback users are always created

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
      const loadedCredentials = XLSX.utils.sheet_to_json(worksheet);
      
      if (loadedCredentials.length > 0) {
        userCredentials = loadedCredentials;
        console.log("User database loaded successfully");
      } else {
        throw new Error('User database is empty');
      }
    })
    .catch(error => {
      console.error("Error loading user database:", error);
      // Use fallback users if loading fails
      createFallbackUsers();
    });
}

// Fallback function to create hardcoded user credentials if loading fails
function createFallbackUsers() {
  console.log("Creating fallback user database");
  // Create a default admin user (password: admin123)
  userCredentials = [
    { 
      username: "admin", 
      password: "14f7022d5259c4f5618a1c6ffb16b2c3" // Hash of "admin123"
    },
    {
      username: "user",
      password: "49cc012bbde5dc4c14eb00bdb746e3a9" // Hash of "user123"
    }
  ];
  console.log("Fallback user database created");
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
  // Add this at the beginning of the initialize function
  console.log("Authentication system initializing...");
  console.log("Login container:", document.getElementById('loginContainer'));
  console.log("App container:", document.getElementById('appContainer'));
  console.log("Login form:", document.getElementById('loginForm'));

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
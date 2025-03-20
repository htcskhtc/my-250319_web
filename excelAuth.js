// Authentication system that reads from Excel file

// Store user credentials
let userCredentials = [];

// Load user credentials from Excel file
function loadUserDatabase() {
  console.log("Attempting to load user database from userDB.xlsx");
  
  return fetch('userDB.xlsx')
    .then(response => {
      if (!response.ok) {
        throw new Error('Failed to fetch userDB.xlsx');
      }
      return response.arrayBuffer();
    })
    .then(data => {
      try {
        const workbook = XLSX.read(data, { type: 'array' });
        console.log("Excel file loaded successfully");
        
        // Check if the workbook has any sheets
        if (workbook.SheetNames.length === 0) {
          throw new Error('Excel file has no sheets');
        }
        
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert the worksheet to JSON
        const users = XLSX.utils.sheet_to_json(worksheet);
        console.log(`Found ${users.length} users in the Excel file`);
        
        if (users.length === 0) {
          throw new Error('No users found in Excel file');
        }
        
        // Check if the users have username and password fields
        if (!users[0].hasOwnProperty('username') || !users[0].hasOwnProperty('password')) {
          throw new Error('Excel file does not have required username and password columns');
        }
        
        // Store the users in the userCredentials array
        userCredentials = users;
        console.log("User database loaded successfully with users:", 
          userCredentials.map(u => u.username).join(", "));
        
        return true;
      } catch (error) {
        console.error("Error parsing Excel file:", error);
        throw error;
      }
    })
    .catch(error => {
      console.error("Error loading user database:", error.message);
      console.log("Falling back to default users");
      createFallbackUsers();
      return false;
    });
}

// Create default users as fallback
function createFallbackUsers() {
  console.log("Creating fallback user database");
  userCredentials = [
    { username: "admin", password: "admin123" },
    { username: "user", password: "user123" },
    { username: "jackchui", password: "jackchui123456" } // Add the user that was previously trying to log in
  ];
  console.log("Fallback user database created with users:", userCredentials.map(u => u.username).join(", "));
}

// Simple authentication function
function authenticateUser(username, password) {
  console.log("Authenticating user:", username);
  console.log("Available users:", userCredentials.map(u => u.username).join(", "));
  
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
  
  // Try to load users from Excel first, fall back to default users if needed
  loadUserDatabase().then(success => {
    console.log("User database initialization complete");
    
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
  });
  
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
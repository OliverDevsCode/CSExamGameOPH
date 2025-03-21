// auth.js
import { PublicClientApplication } from '@azure/msal-browser';

// MSAL configuration object
const msalConfig = {
    auth: {
        clientId: "737b9d9f-9b7b-4dd2-b7d3-aef85aef5565", // Your Application (client) ID from Entra ID
        authority: "https://login.microsoftonline.com/737b9d9f-9b7b-4dd2-b7d3-aef85aef5565", // Tenant-specific endpotint
        redirectUri: "https://cs-exam-game-oph.vercel.app" // Must match your registered redirect URI
    },
    cache: {
        cacheLocation: "sessionStorage", // Options: "sessionStorage" or "localStorage"
        storeAuthStateInCookie: false // Set to true for legacy browsers if needed
    }
};

// Create an instance of PublicClientApplication
const msalInstance = new PublicClientApplication(msalConfig);

// Define a login request
const loginRequest = {
    scopes: ["User.Read"] // Permission to read basic profile information
};

// A function to trigger login using a popup
export function login() {
    msalInstance.loginPopup(loginRequest)
        .then(loginResponse => {
            console.log("Login successful!", loginResponse);
            const userName = loginResponse.account.name;
            // Update the UI if the element exists
            const userInfoEl = document.getElementById("user-info");
            if (userInfoEl) {
                userInfoEl.innerText = `Hello, ${userName}`;
            }
            // Store the user's name in localStorage for use on other pages
            localStorage.setItem("msal_userName", userName);
        })
        .catch(error => {
            console.error("Login failed:", error);
        });
}

export { msalInstance };
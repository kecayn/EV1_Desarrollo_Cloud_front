

export const msalConfig = {
    auth: {
        clientId: "8255d6c9-011f-438c-946d-351cece30dce",
        authority: "https://login.microsoftonline.com/54770a96-6168-4dc5-adb5-87a188b5f997",
        redirectUri: "http://localhost:3000",
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    }
};

export const loginRequest = {
    scopes: ["api://8311e240-9d64-48d5-84ae-cb28494c5414/MesaTech.Access"]
};
package com.NAtools.util;

import com.aspose.email.EWSClient;
import com.aspose.email.EmailClient;
import com.aspose.email.IEWSClient;
import com.aspose.email.OAuthNetworkCredential;
import com.aspose.email.system.ICredentials;
import com.github.scribejava.apis.MicrosoftAzureActiveDirectory20Api;
import com.github.scribejava.core.builder.ServiceBuilder;
import com.github.scribejava.core.model.OAuth2AccessToken;
import com.github.scribejava.core.model.OAuthConstants;
import com.github.scribejava.core.oauth.OAuth20Service;

import java.awt.Desktop;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.Scanner;
import java.util.concurrent.ExecutionException;
import java.util.logging.Logger;

public class OfficeConnectionUtility {

    private static final Logger logger = Logger.getLogger(OfficeConnectionUtility.class.getName());
    private static final String CLIENT_ID = "your_client_id";
    private static final String CLIENT_SECRET = "your_client_secret";
    private static final String AUTHORIZATION_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
    private static final String TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
    private static final String CALLBACK_URL = "http://localhost:8080/callback"; // Adjust as necessary
    private static final String SCOPES = "openid profile offline_access https://outlook.office365.com/EWS.AccessAsUser.All";

    private static OAuth20Service service;
    private static OAuth2AccessToken accessToken;

    // Initialize OAuth2 Service
    private static void initService() {
        service = new ServiceBuilder(CLIENT_ID)
                .apiSecret(CLIENT_SECRET)
                .defaultScope(SCOPES)
                .callback(CALLBACK_URL)
                .build(MicrosoftAzureActiveDirectory20Api.instance());
    }

    public static boolean authenticate() {
        initService();
        try {
            String authorizationUrl = service.getAuthorizationUrl();

            // Open the browser for user authorization
            if (Desktop.isDesktopSupported()) {
                Desktop.getDesktop().browse(new URI(authorizationUrl));
            } else {
                logger.info("Please open the following URL in your browser to authorize:");
                logger.info(authorizationUrl);
            }

            // Get the authorization code from the user
            Scanner in = new Scanner(System.in);
            System.out.println("Enter the authorization code:");
            String authCode = in.nextLine();

            // Exchange authorization code for access token
            accessToken = service.getAccessToken(authCode);
            logger.info("OAuth2 authorization granted. Access Token obtained.");
            return true;

        } catch (IOException | URISyntaxException | InterruptedException | ExecutionException e) {
            logger.severe("Error during OAuth2 authentication: " + e.getMessage());
            e.printStackTrace();
            return false;
        }
    }

    public static IEWSClient connectToOffice365() throws Exception {
        if (accessToken == null) {
            boolean isAuthenticated = authenticate();
            if (!isAuthenticated) {
                throw new Exception("OAuth2 authorization failed. Cannot connect to Office 365.");
            }
        }

        try {
            OAuthNetworkCredential oAuthNetworkCredential = new OAuthNetworkCredential(accessToken.getAccessToken());
            IEWSClient client = EWSClient.getEWSClient("https://outlook.office365.com/ews/exchange.asmx", (ICredentials) oAuthNetworkCredential);
            client.setTimeout(180000);
            EmailClient.setSocketsLayerVersion2(true);
            EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
            logger.info("Connected to Office 365 using OAuth2.");
            return client;

        } catch (Exception e) {
            logger.severe("Failed to connect to Office 365 using OAuth2: " + e.getMessage());
            e.printStackTrace();
            throw new Exception("Failed to connect to Office 365 using OAuth2.");
        }
    }

    public static IEWSClient refreshConnection() throws Exception {
        try {
            String refreshToken = accessToken.getRefreshToken();
            accessToken = service.refreshAccessToken(refreshToken);
            logger.info("Access token refreshed successfully.");

            return connectToOffice365();

        } catch (Exception e) {
            logger.severe("Failed to refresh access token: " + e.getMessage());
            e.printStackTrace();
            throw new Exception("Failed to refresh access token.");
        }
    }
}

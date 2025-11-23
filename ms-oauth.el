;;; ms-oauth.el --- Microsoft OAuth2 PKCE library for Emacs -*- lexical-binding: t -*-

;; Copyright (C) 2025 Ian FitzPatrick

;; Author: Ian FitzPatrick <ian@ianfitzpatrick.eu>
;; URL: github.com/ifitzpat/ms-oauth
;; Version: 0.1.0
;; Package-Requires: ((emacs "27.1") (request "0.3.3") (web-server "0.1.2") (plstore "1.0"))
;; Keywords: oauth microsoft azure authentication

;; This file is not part of GNU Emacs

;; This file is free software; you can redistribute it and/or modify
;; it under the terms of the GNU General Public License as published by
;; the Free Software Foundation; either version 3, or (at your option)
;; any later version.

;; This program is distributed in the hope that it will be useful,
;; but WITHOUT ANY WARRANTY; without even the implied warranty of
;; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
;; GNU General Public License for more details.

;; For a full copy of the GNU General Public License
;; see <https://www.gnu.org/licenses/>.

;;; Commentary:

;; ms-oauth provides a configurable OAuth2 PKCE (Proof Key for Code Exchange)
;; implementation for Microsoft Azure/Graph API authentication.
;;
;; This library allows multiple projects to use different OAuth configurations
;; with secure token storage and automatic token refresh.
;;
;; Usage:
;;
;; 1. Create a configuration:
;;    (setq my-config (ms-oauth-create-config
;;                     :client-id "your-client-id"
;;                     :tenant-id "organizations" 
;;                     :resource-url "https://graph.microsoft.com/Calendars.ReadWrite"
;;                     :token-cache-file "~/.cache/my-tokens.plist"
;;                     :gpg-recipient "user@example.com"))
;;
;; 2. Authenticate:
;;    (ms-oauth-authenticate my-config)
;;
;; 3. Get bearer token for API calls:
;;    (ms-oauth-get-bearer-token my-config)

;;; Code:

(require 'request)
(require 'web-server)
(require 'plstore)
(require 'url-util)
(require 'cl-lib)

;;; Configuration Structure

(cl-defstruct (ms-oauth-config (:constructor ms-oauth--make-config))
  "Configuration structure for Microsoft OAuth2 authentication.

Slots:
- client-id: Azure app registration client ID (required)
- client-secret: Azure app registration client secret (optional for PKCE)
- tenant-id: Azure tenant ID (defaults to \"organizations\")
- resource-url: OAuth scope/resource URL (required)
- token-cache-file: Path to encrypted token cache (required)
- gpg-recipient: GPG key for token encryption (required)
- redirect-uri: OAuth redirect URI (defaults to \"http://localhost:9004\")
- server-port: Local server port for OAuth callback (defaults to 9004)
- auth-url: OAuth authorization URL (auto-generated)
- token-url: OAuth token URL (auto-generated)
- code-verifier: PKCE code verifier (auto-generated)
- code-challenge: PKCE code challenge (auto-generated)"
  client-id
  client-secret
  tenant-id
  resource-url
  token-cache-file
  gpg-recipient
  redirect-uri
  server-port
  auth-url
  token-url
  code-verifier
  code-challenge)

;;; Configuration Management

(defun ms-oauth-create-config (&rest args)
  "Create a new OAuth configuration with the given parameters.

Required parameters:
- :client-id - Azure app registration client ID
- :resource-url - OAuth scope/resource URL  
- :token-cache-file - Path to encrypted token cache
- :gpg-recipient - GPG key for token encryption

Optional parameters:
- :client-secret - Azure app registration client secret
- :tenant-id - Azure tenant ID (defaults to \"organizations\")
- :redirect-uri - OAuth redirect URI (defaults to \"http://localhost:9004\")
- :server-port - Local server port (defaults to 9004)

Example:
  (ms-oauth-create-config
   :client-id \"your-client-id\"
   :resource-url \"https://graph.microsoft.com/Calendars.ReadWrite\"
   :token-cache-file \"~/.cache/my-tokens.plist\"
   :gpg-recipient \"user@example.com\")"
  (let* ((client-id (plist-get args :client-id))
         (client-secret (plist-get args :client-secret))
         (tenant-id (or (plist-get args :tenant-id) "organizations"))
         (resource-url (plist-get args :resource-url))
         (token-cache-file (plist-get args :token-cache-file))
         (gpg-recipient (plist-get args :gpg-recipient))
         (redirect-uri (or (plist-get args :redirect-uri) "http://localhost:9004"))
         (server-port (or (plist-get args :server-port) 9004))
         (auth-url (format "https://login.microsoftonline.com/%s/oauth2/v2.0/authorize" tenant-id))
         (token-url (format "https://login.microsoftonline.com/%s/oauth2/v2.0/token" tenant-id))
         (code-verifier (ms-oauth--generate-random-string 43))
         (code-challenge (base64url-encode-string (secure-hash 'sha256 code-verifier nil nil t))))
    
    ;; Validate required parameters
    (unless client-id
      (error "ms-oauth: client-id is required"))
    (unless resource-url
      (error "ms-oauth: resource-url is required"))
    (unless token-cache-file
      (error "ms-oauth: token-cache-file is required"))
    (unless gpg-recipient
      (error "ms-oauth: gpg-recipient is required"))
    
    ;; Create and return configuration
    (ms-oauth--make-config
     :client-id client-id
     :client-secret client-secret
     :tenant-id tenant-id
     :resource-url resource-url
     :token-cache-file (expand-file-name token-cache-file)
     :gpg-recipient gpg-recipient
     :redirect-uri redirect-uri
     :server-port server-port
     :auth-url auth-url
     :token-url token-url
     :code-verifier code-verifier
     :code-challenge code-challenge)))

(defun ms-oauth-validate-config (config)
  "Validate that CONFIG has all required fields populated.
Returns t if valid, raises error if invalid."
  (unless (ms-oauth-config-p config)
    (error "ms-oauth: Invalid configuration object"))
  
  (let ((required-fields '((client-id . "client-id")
                           (resource-url . "resource-url")
                           (token-cache-file . "token-cache-file")
                           (gpg-recipient . "gpg-recipient"))))
    (dolist (field required-fields)
      (let ((value (funcall (intern (format "ms-oauth-config-%s" (car field))) config))
            (name (cdr field)))
        (unless (and value (not (string-empty-p (if (stringp value) value (format "%s" value)))))
          (error "ms-oauth: Required field %s is missing or empty" name)))))
  t)

;;; Utility Functions

(defun ms-oauth--generate-random-string (length)
  "Generate a random string of LENGTH consisting of URL-safe characters."
  (let ((chars "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~")
        (result ""))
    (dotimes (_ length)
      (setq result (concat result (string (aref chars (random (length chars)))))))
    result))

;;; Error Handling

(define-error 'ms-oauth-error "Microsoft OAuth error")
(define-error 'ms-oauth-config-error "Microsoft OAuth configuration error" 'ms-oauth-error)
(define-error 'ms-oauth-auth-error "Microsoft OAuth authentication error" 'ms-oauth-error)
(define-error 'ms-oauth-token-error "Microsoft OAuth token error" 'ms-oauth-error)

(defun ms-oauth--handle-config-error (config operation)
  "Handle configuration validation errors for CONFIG during OPERATION."
  (condition-case err
      (ms-oauth-validate-config config)
    (error
     (signal 'ms-oauth-config-error 
             (list (format "Invalid configuration for %s: %s" operation (error-message-string err)))))))

;;; Token Management

(defvar ms-oauth--auth-servers (make-hash-table :test 'equal)
  "Hash table storing active auth servers by config.")

(defvar ms-oauth--auth-completion (make-hash-table :test 'equal)
  "Hash table tracking auth completion status by config.")

(defvar ms-oauth--auth-timers (make-hash-table :test 'equal)
  "Hash table storing auth timeout timers by config.")

(defun ms-oauth--token-timed-out (config token-type)
  "Check if TOKEN-TYPE token has timed out for CONFIG.
TOKEN-TYPE should be 'access' or 'refresh'."
  (ms-oauth--handle-config-error config "token timeout check")
  (let* ((cache (plstore-open (ms-oauth-config-token-cache-file config)))
         (cache-key (format "%s-%s" (ms-oauth-config-client-id config) token-type))
         (auth-timestamp (plist-get (cdr (plstore-get cache cache-key)) :timestamp)))
    (plstore-close cache)
    (if auth-timestamp
        (time-less-p (time-add (parse-iso8601-time-string auth-timestamp) 
                               (seconds-to-time 3599))  
                     (current-time))
      nil)))

(defun ms-oauth--token-cache-exists (config)
  "Check if token cache file exists for CONFIG."
  (ms-oauth--handle-config-error config "token cache check")
  (file-regular-p (ms-oauth-config-token-cache-file config)))

(defun ms-oauth--set-token-field (config type public secret)
  "Store token information for CONFIG.
TYPE is the token type (e.g., 'access', 'refresh', 'auth').
PUBLIC is the public metadata.
SECRET is the encrypted token data."
  (ms-oauth--handle-config-error config "token storage")
  (let* ((cache (plstore-open (ms-oauth-config-token-cache-file config)))
         (cache-key (format "%s-%s" (ms-oauth-config-client-id config) type))
         (plstore-encrypt-to (ms-oauth-config-gpg-recipient config)))
    (plstore-put cache cache-key public secret)
    (plstore-save cache)
    (plstore-close cache)))

(defun ms-oauth--get-auth-token (config)
  "Retrieve auth token for CONFIG."
  (ms-oauth--handle-config-error config "auth token retrieval")
  (let* ((cache (plstore-open (ms-oauth-config-token-cache-file config)))
         (cache-key (format "%s-auth" (ms-oauth-config-client-id config)))
         (token (plist-get (cdr (plstore-get cache cache-key)) :auth)))
    (plstore-close cache)
    token))

(defun ms-oauth--get-access-token (config)
  "Retrieve access token for CONFIG."
  (ms-oauth--handle-config-error config "access token retrieval")
  (let* ((cache (plstore-open (ms-oauth-config-token-cache-file config)))
         (cache-key (format "%s-access" (ms-oauth-config-client-id config)))
         (token (plist-get (cdr (plstore-get cache cache-key)) :access)))
    (plstore-close cache)
    token))

(defun ms-oauth--get-refresh-token (config)
  "Retrieve refresh token for CONFIG if it exists and hasn't expired."
  (ms-oauth--handle-config-error config "refresh token retrieval")
  (let* ((cache (plstore-open (ms-oauth-config-token-cache-file config)))
         (cache-key (format "%s-refresh" (ms-oauth-config-client-id config)))
         (token (plist-get (cdr (plstore-get cache cache-key)) :refresh)))
    (plstore-close cache)
    (and token (not (ms-oauth--token-timed-out config "refresh")) token)))

;;; Authorization Flow Functions

(defun ms-oauth--process-token-callback (config request)
  "Process OAuth callback for CONFIG and signal completion."
  (with-slots (process headers) request
    (ws-response-header process 200 '("Content-type" . "text/html"))
    (let ((auth-code (cdr (assoc "code" (headers request))))
          (config-key (ms-oauth-config-client-id config)))
      (if auth-code
          (progn
            (ms-oauth--set-token-field config "auth" nil `(:auth ,auth-code))
            (puthash config-key t ms-oauth--auth-completion)
            (let ((timer (gethash config-key ms-oauth--auth-timers)))
              (when timer
                (cancel-timer timer)))
            (process-send-string process
                                 "<html><body><h2>Authentication successful!</h2><p>You can close this window and return to Emacs.</p></body></html>"))
        (process-send-string process
                             "<html><body><h2>Authentication failed</h2><p>No authorization code received.</p></body></html>")))))

(defun ms-oauth--start-auth-code-server (config)
  "Start OAuth callback server for CONFIG."
  (let ((port (ms-oauth-config-server-port config)))
    (ws-start
     `(((:GET . ".*") .
        (lambda (request)
          (ms-oauth--process-token-callback ,config request))))
     port)))

(defun ms-oauth--request-authorization (config)
  "Request OAuth authorization for CONFIG without blocking Emacs.
Opens the browser and waits for the callback with a 5-minute timeout.
Returns the authorization code on success."
  (ms-oauth--handle-config-error config "authorization request")
  (let* ((config-key (ms-oauth-config-client-id config))
         (auth-url
          (concat (ms-oauth-config-auth-url config)
                  "?client_id=" (url-hexify-string (ms-oauth-config-client-id config))
                  "&response_type=code"
                  "&code_challenge=" (ms-oauth-config-code-challenge config)
                  "&code_challenge_method=S256"
                  "&redirect_uri=" (url-hexify-string (ms-oauth-config-redirect-uri config))
                  "&scope=" (url-hexify-string (concat "offline_access " (ms-oauth-config-resource-url config))))))

    ;; Reset state
    (puthash config-key nil ms-oauth--auth-completion)

    ;; Start server
    (puthash config-key (ms-oauth--start-auth-code-server config) ms-oauth--auth-servers)

    ;; Set timeout (5 minutes)
    (puthash config-key
             (run-at-time 300 nil
                          (lambda ()
                            (unless (gethash config-key ms-oauth--auth-completion)
                              (let ((server (gethash config-key ms-oauth--auth-servers)))
                                (when server
                                  (ws-stop server)))
                              (remhash config-key ms-oauth--auth-servers)
                              (signal 'ms-oauth-auth-error 
                                      '("OAuth authorization timed out after 5 minutes")))))
             ms-oauth--auth-timers)

    ;; Open browser (non-blocking)
    (browse-url-xdg-open auth-url)
    (message "Please complete authentication in your browser (you have 5 minutes)...")

    ;; Wait for completion with non-blocking checks
    (while (not (gethash config-key ms-oauth--auth-completion))
      (accept-process-output nil 0.1))

    ;; Cleanup
    (let ((server (gethash config-key ms-oauth--auth-servers)))
      (when server
        (ws-stop server)))
    (remhash config-key ms-oauth--auth-servers)
    (let ((timer (gethash config-key ms-oauth--auth-timers)))
      (when timer
        (cancel-timer timer)))
    (remhash config-key ms-oauth--auth-timers)

    (message "Authentication successful!")
    (ms-oauth--get-auth-token config)))

(defun ms-oauth--request-access-token (config)
  "Request access token for CONFIG using either refresh token or authorization code."
  (ms-oauth--handle-config-error config "access token request")
  (let* ((refresh-token (ms-oauth--get-refresh-token config))
         (auth-code (or refresh-token
                        (progn
                          (ms-oauth--request-authorization config)))))
    
    (condition-case err
        (request (ms-oauth-config-token-url config)
          :type "POST"
          :sync t
          :data `(("tenant" . ,(ms-oauth-config-tenant-id config))
                  ("client_id" . ,(ms-oauth-config-client-id config))
                  ("scope" . ,(concat "offline_access " (ms-oauth-config-resource-url config)))
                  ,(if refresh-token
                       `("refresh_token" . ,refresh-token)
                     `("code" . ,auth-code))
                  ("redirect_uri" . ,(ms-oauth-config-redirect-uri config))
                  ,(if refresh-token
                       '("grant_type" . "refresh_token")
                     '("grant_type" . "authorization_code"))
                  ("code_verifier" . ,(ms-oauth-config-code-verifier config)))
          :parser 'json-read
          :success (cl-function
                    (lambda (&key data &allow-other-keys)
                      (when data
                        (ms-oauth--set-token-field config "access" 
                                                   `(:timestamp ,(format-time-string "%Y-%m-%dT%H:%M:%S" (current-time))) 
                                                   `(:access ,(assoc-default 'access_token data)))
                        (ms-oauth--set-token-field config "refresh" 
                                                   `(:timestamp ,(format-time-string "%Y-%m-%dT%H:%M:%S" (current-time))) 
                                                   `(:refresh ,(assoc-default 'refresh_token data))))))
          :error (cl-function
                  (lambda (&rest args &key error-thrown &allow-other-keys)
                    (signal 'ms-oauth-token-error 
                            (list (format "Error getting access token: %S" error-thrown))))))
      (error
       (signal 'ms-oauth-token-error 
               (list (format "Failed to request access token: %s" (error-message-string err))))))))

;;; Public API Functions

(defun ms-oauth-get-bearer-token (config)
  "Get a valid bearer token for CONFIG.
Automatically refreshes token if expired or requests new one if none exists."
  (ms-oauth--handle-config-error config "bearer token request")
  (when (or (ms-oauth--token-timed-out config "access")
            (not (ms-oauth--token-cache-exists config)))
    (ms-oauth--request-access-token config))
  (ms-oauth--get-access-token config))

(defun ms-oauth-authenticate (config)
  "Authenticate with Microsoft OAuth using CONFIG.
This is the main entry point for authentication."
  (ms-oauth--handle-config-error config "authentication")
  (ms-oauth--request-access-token config)
  (message "Microsoft OAuth authentication completed successfully"))

(defun ms-oauth-tokens-valid-p (config)
  "Check if stored tokens for CONFIG are valid and not expired."
  (ms-oauth--handle-config-error config "token validation")
  (and (ms-oauth--token-cache-exists config)
       (not (ms-oauth--token-timed-out config "access"))))

(defun ms-oauth-refresh-tokens (config)
  "Force refresh of tokens for CONFIG."
  (ms-oauth--handle-config-error config "token refresh")
  (ms-oauth--request-access-token config)
  (message "Tokens refreshed successfully"))

(defun ms-oauth-clear-tokens (config)
  "Clear stored tokens for CONFIG."
  (ms-oauth--handle-config-error config "token clearing")
  (let ((cache-file (ms-oauth-config-token-cache-file config))
        (client-id (ms-oauth-config-client-id config)))
    (when (file-exists-p cache-file)
      (let ((cache (plstore-open cache-file)))
        (dolist (token-type '("access" "refresh" "auth"))
          (let ((cache-key (format "%s-%s" client-id token-type)))
            (plstore-delete cache cache-key)))
        (plstore-save cache)
        (plstore-close cache)))
    (message "Tokens cleared successfully")))

(defun ms-oauth-revoke-tokens (config)
  "Revoke and clear stored tokens for CONFIG.
Note: Microsoft Graph API doesn't support token revocation endpoint,
so this just clears local tokens."
  (ms-oauth-clear-tokens config))

(provide 'ms-oauth)

;;; ms-oauth.el ends here
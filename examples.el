;;; examples.el --- Usage examples for ms-oauth library -*- lexical-binding: t -*-

;; Copyright (C) 2025 Ian FitzPatrick

;; Author: Ian FitzPatrick <ian@ianfitzpatrick.eu>

;; This file is not part of GNU Emacs

;;; Commentary:

;; This file provides practical examples of how to use the ms-oauth library
;; for different Microsoft Graph API scenarios.

;;; Code:

(require 'ms-oauth)
(require 'request)

;;; Example 1: Outlook Calendar Integration

(defun example-outlook-calendar-setup ()
  "Example: Set up OAuth for Outlook calendar access."
  ;; Create configuration for calendar access
  (setq outlook-calendar-config
        (ms-oauth-create-config
         :client-id "your-azure-app-client-id"           ; Replace with your client ID
         :tenant-id "organizations"                       ; Or your specific tenant ID
         :resource-url "https://graph.microsoft.com/Calendars.ReadWrite"
         :token-cache-file "~/.cache/outlook-calendar-tokens.plist"
         :gpg-recipient "your-email@example.com"))       ; Replace with your GPG key email
  
  ;; Authenticate (this will open browser for initial auth)
  (ms-oauth-authenticate outlook-calendar-config)
  
  (message "Outlook calendar authentication configured"))

(defun example-outlook-get-calendar-events ()
  "Example: Get calendar events from Outlook."
  (interactive)
  (let ((bearer-token (ms-oauth-get-bearer-token outlook-calendar-config)))
    (request "https://graph.microsoft.com/v1.0/me/calendar/events"
      :headers `(("Authorization" . ,(concat "Bearer " bearer-token))
                 ("Content-Type" . "application/json"))
      :parser 'json-read
      :success (cl-function
                (lambda (&key data &allow-other-keys)
                  (message "Found %d calendar events" 
                           (length (assoc-default 'value data)))
                  (dolist (event (assoc-default 'value data))
                    (message "Event: %s" (assoc-default 'subject event)))))
      :error (cl-function
              (lambda (&rest args &key error-thrown &allow-other-keys)
                (message "Error retrieving calendar events: %S" error-thrown))))))

(defun example-outlook-create-calendar-event (subject start-time end-time)
  "Example: Create a calendar event in Outlook."
  (interactive "sEvent subject: \nsStart time (ISO 8601): \nsEnd time (ISO 8601): ")
  (let ((bearer-token (ms-oauth-get-bearer-token outlook-calendar-config))
        (event-data `(("subject" . ,subject)
                      ("start" . (("dateTime" . ,start-time)
                                  ("timeZone" . "UTC")))
                      ("end" . (("dateTime" . ,end-time)
                                ("timeZone" . "UTC"))))))
    
    (request "https://graph.microsoft.com/v1.0/me/calendar/events"
      :type "POST"
      :headers `(("Authorization" . ,(concat "Bearer " bearer-token))
                 ("Content-Type" . "application/json"))
      :data (json-encode event-data)
      :parser 'json-read
      :success (cl-function
                (lambda (&key data &allow-other-keys)
                  (message "Created event: %s" (assoc-default 'subject data))))
      :error (cl-function
              (lambda (&rest args &key error-thrown &allow-other-keys)
                (message "Error creating event: %S" error-thrown))))))

;;; Example 2: OneDrive File Access

(defun example-onedrive-setup ()
  "Example: Set up OAuth for OneDrive file access."
  (setq onedrive-config
        (ms-oauth-create-config
         :client-id "your-azure-app-client-id"           ; Replace with your client ID
         :tenant-id "organizations"
         :resource-url "https://graph.microsoft.com/Files.ReadWrite"
         :token-cache-file "~/.cache/onedrive-tokens.plist"
         :gpg-recipient "your-email@example.com"))       ; Replace with your GPG key email
  
  (ms-oauth-authenticate onedrive-config)
  (message "OneDrive authentication configured"))

(defun example-onedrive-list-files ()
  "Example: List files in OneDrive root folder."
  (interactive)
  (let ((bearer-token (ms-oauth-get-bearer-token onedrive-config)))
    (request "https://graph.microsoft.com/v1.0/me/drive/root/children"
      :headers `(("Authorization" . ,(concat "Bearer " bearer-token)))
      :parser 'json-read
      :success (cl-function
                (lambda (&key data &allow-other-keys)
                  (message "Found %d files in OneDrive" 
                           (length (assoc-default 'value data)))
                  (dolist (file (assoc-default 'value data))
                    (message "File: %s (Size: %s bytes)" 
                             (assoc-default 'name file)
                             (assoc-default 'size file)))))
      :error (cl-function
              (lambda (&rest args &key error-thrown &allow-other-keys)
                (message "Error listing OneDrive files: %S" error-thrown))))))

(defun example-onedrive-upload-file (local-file-path remote-filename)
  "Example: Upload a file to OneDrive."
  (interactive "fLocal file: \nsRemote filename: ")
  (let ((bearer-token (ms-oauth-get-bearer-token onedrive-config))
        (file-content (with-temp-buffer
                        (insert-file-contents local-file-path)
                        (buffer-string))))
    
    (request (format "https://graph.microsoft.com/v1.0/me/drive/root:/%s:/content" 
                     (url-hexify-string remote-filename))
      :type "PUT"
      :headers `(("Authorization" . ,(concat "Bearer " bearer-token))
                 ("Content-Type" . "application/octet-stream"))
      :data file-content
      :parser 'json-read
      :success (cl-function
                (lambda (&key data &allow-other-keys)
                  (message "Uploaded file: %s" (assoc-default 'name data))))
      :error (cl-function
              (lambda (&rest args &key error-thrown &allow-other-keys)
                (message "Error uploading file: %S" error-thrown))))))

;;; Example 3: Microsoft Teams Integration

(defun example-teams-setup ()
  "Example: Set up OAuth for Microsoft Teams access."
  (setq teams-config
        (ms-oauth-create-config
         :client-id "your-azure-app-client-id"           ; Replace with your client ID
         :tenant-id "organizations"
         :resource-url "https://graph.microsoft.com/Team.ReadBasic.All"
         :token-cache-file "~/.cache/teams-tokens.plist"
         :gpg-recipient "your-email@example.com"))       ; Replace with your GPG key email
  
  (ms-oauth-authenticate teams-config)
  (message "Teams authentication configured"))

(defun example-teams-list-joined-teams ()
  "Example: List Teams that the user has joined."
  (interactive)
  (let ((bearer-token (ms-oauth-get-bearer-token teams-config)))
    (request "https://graph.microsoft.com/v1.0/me/joinedTeams"
      :headers `(("Authorization" . ,(concat "Bearer " bearer-token)))
      :parser 'json-read
      :success (cl-function
                (lambda (&key data &allow-other-keys)
                  (message "Found %d Teams" 
                           (length (assoc-default 'value data)))
                  (dolist (team (assoc-default 'value data))
                    (message "Team: %s (%s)" 
                             (assoc-default 'displayName team)
                             (assoc-default 'description team)))))
      :error (cl-function
              (lambda (&rest args &key error-thrown &allow-other-keys)
                (message "Error listing Teams: %S" error-thrown))))))

;;; Example 4: Multiple Configuration Management

(defvar ms-oauth-example-configs (make-hash-table :test 'equal)
  "Hash table to store multiple OAuth configurations.")

(defun example-setup-multiple-configs ()
  "Example: Set up multiple OAuth configurations for different services."
  ;; Calendar configuration
  (puthash "calendar"
           (ms-oauth-create-config
            :client-id "calendar-client-id"
            :resource-url "https://graph.microsoft.com/Calendars.ReadWrite"
            :token-cache-file "~/.cache/calendar-tokens.plist"
            :gpg-recipient "user@example.com")
           ms-oauth-example-configs)
  
  ;; OneDrive configuration  
  (puthash "onedrive"
           (ms-oauth-create-config
            :client-id "onedrive-client-id"
            :resource-url "https://graph.microsoft.com/Files.ReadWrite"
            :token-cache-file "~/.cache/onedrive-tokens.plist"
            :gpg-recipient "user@example.com")
           ms-oauth-example-configs)
  
  ;; Teams configuration
  (puthash "teams"
           (ms-oauth-create-config
            :client-id "teams-client-id"
            :resource-url "https://graph.microsoft.com/Team.ReadBasic.All"
            :token-cache-file "~/.cache/teams-tokens.plist"
            :gpg-recipient "user@example.com")
           ms-oauth-example-configs)
  
  (message "Multiple OAuth configurations set up"))

(defun example-authenticate-all-configs ()
  "Example: Authenticate all configured services."
  (interactive)
  (maphash (lambda (service config)
             (message "Authenticating %s..." service)
             (ms-oauth-authenticate config)
             (message "%s authentication complete" service))
           ms-oauth-example-configs))

(defun example-get-bearer-token-for-service (service)
  "Example: Get bearer token for a specific service."
  (let ((config (gethash service ms-oauth-example-configs)))
    (if config
        (ms-oauth-get-bearer-token config)
      (error "Service %s not configured" service))))

;;; Example 5: Error Handling and Token Management

(defun example-robust-api-call (config url)
  "Example: Make a robust API call with error handling and token refresh."
  (condition-case err
      (let ((bearer-token (ms-oauth-get-bearer-token config)))
        (request url
          :headers `(("Authorization" . ,(concat "Bearer " bearer-token)))
          :parser 'json-read
          :success (cl-function
                    (lambda (&key data &allow-other-keys)
                      (message "API call successful: %S" data)))
          :error (cl-function
                  (lambda (&rest args &key response &allow-other-keys)
                    (let ((status-code (request-response-status-code response)))
                      (cond 
                       ((= status-code 401)
                        (message "Token expired, refreshing...")
                        (ms-oauth-refresh-tokens config)
                        ;; Retry the call (implement retry logic as needed)
                        (message "Please retry the operation"))
                       (t
                        (message "API call failed with status %d" status-code))))))))
    (ms-oauth-token-error
     (message "Token error: %s" (error-message-string err))
     (message "Try clearing tokens: (ms-oauth-clear-tokens config)"))
    (ms-oauth-auth-error
     (message "Authentication error: %s" (error-message-string err)))
    (error
     (message "Unexpected error: %s" (error-message-string err)))))

(defun example-check-token-status (config)
  "Example: Check and display token status."
  (interactive)
  (cond
   ((ms-oauth-tokens-valid-p config)
    (message "Tokens are valid and current"))
   ((ms-oauth--token-cache-exists config)
    (message "Tokens exist but may be expired - refresh needed"))
   (t
    (message "No tokens found - authentication needed"))))

;;; Example 6: Configuration Validation

(defun example-validate-configurations ()
  "Example: Validate all stored configurations."
  (interactive)
  (maphash (lambda (service config)
             (condition-case err
                 (progn
                   (ms-oauth-validate-config config)
                   (message "%s configuration is valid" service))
               (ms-oauth-config-error
                (message "%s configuration is invalid: %s" 
                         service (error-message-string err)))))
           ms-oauth-example-configs))

;;; Running the Examples

(defun ms-oauth-run-calendar-example ()
  "Run the Outlook calendar example."
  (interactive)
  (example-outlook-calendar-setup)
  (example-outlook-get-calendar-events))

(defun ms-oauth-run-onedrive-example ()
  "Run the OneDrive example."
  (interactive)
  (example-onedrive-setup)
  (example-onedrive-list-files))

(defun ms-oauth-run-teams-example ()
  "Run the Teams example."
  (interactive)
  (example-teams-setup)
  (example-teams-list-joined-teams))

(provide 'examples)

;;; examples.el ends here
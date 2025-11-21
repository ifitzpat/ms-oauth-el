;;; test-ms-oauth.el --- Test suite for ms-oauth library -*- lexical-binding: t -*-

;; Copyright (C) 2025 Ian FitzPatrick

;; Author: Ian FitzPatrick <ian@ianfitzpatrick.eu>

;; This file is not part of GNU Emacs

;;; Commentary:

;; Test suite for the ms-oauth library.
;; Provides unit tests and integration tests for OAuth functionality.

;;; Code:

(require 'ert)
(require 'package)
(add-to-list 'package-archives '("melpa" . "https://melpa.org/packages/"))
(package-initialize)
(add-to-list 'load-path ".")

(require 'ms-oauth)

;;; Test Configuration Creation

(ert-deftest ms-oauth-test-create-config-valid ()
  "Test creating a valid configuration."
  (let ((config (ms-oauth-create-config
                 :client-id "test-client-id"
                 :resource-url "https://graph.microsoft.com/Calendars.ReadWrite"
                 :token-cache-file "/tmp/test-tokens.plist"
                 :gpg-recipient "test@example.com")))
    (should (ms-oauth-config-p config))
    (should (string= "test-client-id" (ms-oauth-config-client-id config)))
    (should (string= "organizations" (ms-oauth-config-tenant-id config)))
    (should (string= "https://graph.microsoft.com/Calendars.ReadWrite" 
                     (ms-oauth-config-resource-url config)))
    (should (string= "/tmp/test-tokens.plist" (ms-oauth-config-token-cache-file config)))
    (should (string= "test@example.com" (ms-oauth-config-gpg-recipient config)))
    (should (string= "http://localhost:9004" (ms-oauth-config-redirect-uri config)))
    (should (= 9004 (ms-oauth-config-server-port config)))))

(ert-deftest ms-oauth-test-create-config-custom-values ()
  "Test creating configuration with custom values."
  (let ((config (ms-oauth-create-config
                 :client-id "custom-client"
                 :tenant-id "custom-tenant"
                 :resource-url "https://graph.microsoft.com/Files.ReadWrite"
                 :token-cache-file "/custom/path/tokens.plist"
                 :gpg-recipient "custom@example.com"
                 :redirect-uri "http://localhost:8080"
                 :server-port 8080)))
    (should (string= "custom-tenant" (ms-oauth-config-tenant-id config)))
    (should (string= "http://localhost:8080" (ms-oauth-config-redirect-uri config)))
    (should (= 8080 (ms-oauth-config-server-port config)))))

(ert-deftest ms-oauth-test-create-config-missing-client-id ()
  "Test that missing client-id raises an error."
  (should-error 
   (ms-oauth-create-config
    :resource-url "https://graph.microsoft.com/Calendars.ReadWrite"
    :token-cache-file "/tmp/test-tokens.plist"
    :gpg-recipient "test@example.com")
   :type 'error))

(ert-deftest ms-oauth-test-create-config-missing-resource-url ()
  "Test that missing resource-url raises an error."
  (should-error
   (ms-oauth-create-config
    :client-id "test-client-id"
    :token-cache-file "/tmp/test-tokens.plist"
    :gpg-recipient "test@example.com")
   :type 'error))

(ert-deftest ms-oauth-test-create-config-missing-token-cache-file ()
  "Test that missing token-cache-file raises an error."
  (should-error
   (ms-oauth-create-config
    :client-id "test-client-id"
    :resource-url "https://graph.microsoft.com/Calendars.ReadWrite"
    :gpg-recipient "test@example.com")
   :type 'error))

(ert-deftest ms-oauth-test-create-config-missing-gpg-recipient ()
  "Test that missing gpg-recipient raises an error."
  (should-error
   (ms-oauth-create-config
    :client-id "test-client-id"
    :resource-url "https://graph.microsoft.com/Calendars.ReadWrite"
    :token-cache-file "/tmp/test-tokens.plist")
   :type 'error))

;;; Test Configuration Validation

(ert-deftest ms-oauth-test-validate-config-valid ()
  "Test validation of valid configuration."
  (let ((config (ms-oauth-create-config
                 :client-id "test-client-id"
                 :resource-url "https://graph.microsoft.com/Calendars.ReadWrite"
                 :token-cache-file "/tmp/test-tokens.plist"
                 :gpg-recipient "test@example.com")))
    (should (ms-oauth-validate-config config))))

(ert-deftest ms-oauth-test-validate-config-invalid-object ()
  "Test validation with invalid configuration object."
  (should-error
   (ms-oauth-validate-config "not-a-config")
   :type 'error))

(ert-deftest ms-oauth-test-validate-config-nil ()
  "Test validation with nil configuration."
  (should-error
   (ms-oauth-validate-config nil)
   :type 'error))

;;; Test PKCE Generation

(ert-deftest ms-oauth-test-pkce-generation ()
  "Test that PKCE values are generated correctly."
  (let ((config (ms-oauth-create-config
                 :client-id "test-client-id"
                 :resource-url "https://graph.microsoft.com/Calendars.ReadWrite"
                 :token-cache-file "/tmp/test-tokens.plist"
                 :gpg-recipient "test@example.com")))
    (should (ms-oauth-config-code-verifier config))
    (should (ms-oauth-config-code-challenge config))
    (should (= 43 (length (ms-oauth-config-code-verifier config))))
    ;; Code challenge should be different from verifier (it's hashed and encoded)
    (should (not (string= (ms-oauth-config-code-verifier config)
                          (ms-oauth-config-code-challenge config))))))

(ert-deftest ms-oauth-test-unique-pkce-values ()
  "Test that each configuration gets unique PKCE values."
  (let ((config1 (ms-oauth-create-config
                  :client-id "test-client-1"
                  :resource-url "https://graph.microsoft.com/Calendars.ReadWrite"
                  :token-cache-file "/tmp/test-tokens1.plist"
                  :gpg-recipient "test@example.com"))
        (config2 (ms-oauth-create-config
                  :client-id "test-client-2"
                  :resource-url "https://graph.microsoft.com/Files.ReadWrite"
                  :token-cache-file "/tmp/test-tokens2.plist"
                  :gpg-recipient "test@example.com")))
    (should (not (string= (ms-oauth-config-code-verifier config1)
                          (ms-oauth-config-code-verifier config2))))
    (should (not (string= (ms-oauth-config-code-challenge config1)
                          (ms-oauth-config-code-challenge config2))))))

;;; Test URL Generation

(ert-deftest ms-oauth-test-url-generation ()
  "Test that OAuth URLs are generated correctly."
  (let ((config (ms-oauth-create-config
                 :client-id "test-client-id"
                 :tenant-id "test-tenant"
                 :resource-url "https://graph.microsoft.com/Calendars.ReadWrite"
                 :token-cache-file "/tmp/test-tokens.plist"
                 :gpg-recipient "test@example.com")))
    (should (string= "https://login.microsoftonline.com/test-tenant/oauth2/v2.0/authorize"
                     (ms-oauth-config-auth-url config)))
    (should (string= "https://login.microsoftonline.com/test-tenant/oauth2/v2.0/token"
                     (ms-oauth-config-token-url config)))))

;;; Test Error Handling

(ert-deftest ms-oauth-test-error-handling-bearer-token ()
  "Test error handling for bearer token with invalid config."
  (should-error
   (ms-oauth-get-bearer-token nil)
   :type 'ms-oauth-config-error))

(ert-deftest ms-oauth-test-error-handling-authenticate ()
  "Test error handling for authenticate with invalid config."
  (should-error
   (ms-oauth-authenticate "invalid-config")
   :type 'ms-oauth-config-error))

(ert-deftest ms-oauth-test-error-handling-tokens-valid-p ()
  "Test error handling for tokens-valid-p with invalid config."
  (should-error
   (ms-oauth-tokens-valid-p nil)
   :type 'ms-oauth-config-error))

;;; Integration Test Helpers

(defun ms-oauth-test-create-test-config ()
  "Create a test configuration for integration tests."
  (ms-oauth-create-config
   :client-id "test-client-id"
   :resource-url "https://graph.microsoft.com/Calendars.ReadWrite"
   :token-cache-file (make-temp-file "ms-oauth-test-" nil ".plist")
   :gpg-recipient "test@example.com"
   :server-port 9999)) ; Use different port to avoid conflicts

(ert-deftest ms-oauth-test-integration-config-creation ()
  "Integration test for creating and using a test configuration."
  (let ((config (ms-oauth-test-create-test-config)))
    (should (ms-oauth-config-p config))
    (should (ms-oauth-validate-config config))
    ;; Clean up temp file
    (when (file-exists-p (ms-oauth-config-token-cache-file config))
      (delete-file (ms-oauth-config-token-cache-file config)))))

;;; Test Runner

(defun ms-oauth-run-tests ()
  "Run all ms-oauth tests."
  (interactive)
  (ert "ms-oauth-test-.*"))

(provide 'test-ms-oauth)

;;; test-ms-oauth.el ends here
# Makefile for ms-oauth library

.PHONY: test install-deps clean help lint

# Default target
all: test

# Install required dependencies
install-deps:
	@echo "Installing dependencies..."
	@emacs --batch --eval "\
	(require 'package) \
	(add-to-list 'package-archives '(\"melpa\" . \"https://melpa.org/packages/\")) \
	(package-initialize) \
	(unless (package-installed-p 'request) \
	  (package-refresh-contents) \
	  (package-install 'request)) \
	(unless (package-installed-p 'web-server) \
	  (package-refresh-contents) \
	  (package-install 'web-server))"
	@echo "Dependencies installed successfully!"

# Run tests
test: install-deps
	@echo "Running ms-oauth tests..."
	@emacs --batch \
		--eval "(require 'package)" \
		--eval "(add-to-list 'package-archives '(\"melpa\" . \"https://melpa.org/packages/\"))" \
		--eval "(package-initialize)" \
		--eval "(add-to-list 'load-path \".\")" \
		-l test-ms-oauth.el \
		-f ms-oauth-run-tests
	@echo "Tests completed!"

# Basic syntax check
lint:
	@echo "Checking syntax..."
	@emacs --batch \
		--eval "(require 'package)" \
		--eval "(add-to-list 'package-archives '(\"melpa\" . \"https://melpa.org/packages/\"))" \
		--eval "(package-initialize)" \
		--eval "(add-to-list 'load-path \".\")" \
		-l ms-oauth.el \
		--eval "(message \"Syntax check passed\")"
	@echo "Syntax check completed!"

# Clean up temporary files
clean:
	@echo "Cleaning up..."
	@find . -name "*.elc" -delete
	@find . -name "*~" -delete
	@find . -name "#*#" -delete
	@echo "Cleanup completed!"

# Display help
help:
	@echo "Available targets:"
	@echo "  test         - Run the test suite (installs deps first)"
	@echo "  install-deps - Install required Emacs packages"
	@echo "  lint         - Check syntax of ms-oauth.el"
	@echo "  clean        - Remove temporary files"
	@echo "  help         - Show this help message"
#!/bin/bash
set -e

# ---------------------------------------------------------------------------
# Color output helpers
# ---------------------------------------------------------------------------
info()    { printf "\033[0;36m[INFO]\033[0m  %s\n" "$*"; }
success() { printf "\033[0;32m[OK]\033[0m    %s\n" "$*"; }
error()   { printf "\033[0;31m[ERROR]\033[0m %s\n" "$*" >&2; }

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
REPO_URL="https://github.com/SuperAllenYao/auto-excel.git"
INSTALL_DIR="$HOME/.auto-excel"
BIN_DIR="$HOME/.local/bin"
BIN_FILE="$BIN_DIR/auto-excel"

# ---------------------------------------------------------------------------
# Step 1: Clone or update the repository
# ---------------------------------------------------------------------------
if [ -d "$INSTALL_DIR/.git" ]; then
    info "Repository already exists at $INSTALL_DIR — pulling latest changes..."
    git -C "$INSTALL_DIR" pull
    success "Repository updated."
else
    info "Cloning auto-excel to $INSTALL_DIR..."
    git clone "$REPO_URL" "$INSTALL_DIR"
    success "Repository cloned."
fi

# ---------------------------------------------------------------------------
# Step 2: Ensure uv is available
# ---------------------------------------------------------------------------
if ! command -v uv >/dev/null 2>&1; then
    info "uv not found — installing uv..."
    curl -LsSf https://astral.sh/uv/install.sh | sh
    # Make uv available in the current shell session
    export PATH="$HOME/.local/bin:$PATH"
    if ! command -v uv >/dev/null 2>&1; then
        error "uv installation failed or not found in PATH. Please install uv manually and re-run."
        exit 1
    fi
    success "uv installed."
else
    success "uv is already installed: $(command -v uv)"
fi

# ---------------------------------------------------------------------------
# Step 3: Install Python dependencies
# ---------------------------------------------------------------------------
info "Running uv sync in $INSTALL_DIR..."
(cd "$INSTALL_DIR" && uv sync)
success "Dependencies synced."

# ---------------------------------------------------------------------------
# Step 4: Create the wrapper script
# ---------------------------------------------------------------------------
info "Creating wrapper script at $BIN_FILE..."
mkdir -p "$BIN_DIR"
cat > "$BIN_FILE" << 'EOF'
#!/bin/bash
cd ~/.auto-excel && uv run auto-excel "$@"
EOF
chmod +x "$BIN_FILE"
success "Wrapper script created."

# ---------------------------------------------------------------------------
# Step 5: Ensure ~/.local/bin is in PATH
# ---------------------------------------------------------------------------
if [[ ":$PATH:" != *":$BIN_DIR:"* ]]; then
    info "$BIN_DIR is not in PATH — appending to ~/.zshrc..."
    echo '' >> "$HOME/.zshrc"
    echo '# Added by auto-excel installer' >> "$HOME/.zshrc"
    echo 'export PATH="$HOME/.local/bin:$PATH"' >> "$HOME/.zshrc"
    success "PATH updated in ~/.zshrc (restart your shell or run: source ~/.zshrc)."
else
    success "$BIN_DIR is already in PATH."
fi

# ---------------------------------------------------------------------------
# Step 6: Create required working directories
# ---------------------------------------------------------------------------
info "Creating working directories on Desktop..."
mkdir -p "$HOME/Desktop/marketing analysis/Raw" \
         "$HOME/Desktop/marketing analysis/New" \
         "$HOME/Desktop/marketing analysis/log"
success "Working directories ready."

# ---------------------------------------------------------------------------
# Done
# ---------------------------------------------------------------------------
echo ""
success "======================================================"
success " auto-excel installation complete!"
success "======================================================"
echo ""
echo "  To start monitoring and processing Excel files, run:"
echo ""
echo "      auto-excel on"
echo ""
echo "  If 'auto-excel' is not found, restart your shell or run:"
echo ""
echo "      source ~/.zshrc"
echo ""

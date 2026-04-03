# MCP Registry Publishing Guide
# ================================
# Run these commands ONE-BY-ONE (DO NOT run as a script!) from the mcp-server/ directory
# cd mcp-server

# ============================================
# STEP 1: npm Login (already done if npm whoami works)
# ============================================
# npm adduser  # (skip if already logged in)
# npm whoami   # verify

# ============================================
# STEP 2: Publish to npm (using granular access token)
# ============================================
# npm requires 2FA or a granular token for scoped packages.
# Go to: https://www.npmjs.com/settings/tokens/granular-access-tokens/new
# Create a token with:
#   - Token name: markdown-mcp-publish
#   - Expiration: 30 days (or as needed)
#   - Permissions: Read and write
#   - Select packages: @xjtlumedia/markdown-mcp-server  (or All packages)
# Then run:
npm publish --access public --auth-type=web
# Or with token directly:
# $env:NPM_TOKEN="your-granular-token-here"
# npm publish --access public
# Verify at: https://www.npmjs.com/package/@xjtlumedia/markdown-mcp-server

# ============================================
# STEP 3: mcp-publisher is already installed!
# ============================================
# mcp-publisher --help

# ============================================
# STEP 4: server.json is already created and validated!
# ============================================
# mcp-publisher validate

# ============================================
# STEP 5: Authenticate with MCP Registry (GitHub) — needs proxy!
# ============================================
# Set proxy for mcp-publisher (uses your local Clash proxy)
$env:HTTPS_PROXY = "http://127.0.0.1:7890"
$env:HTTP_PROXY = "http://127.0.0.1:7890"
mcp-publisher login github
# Follow the browser prompts with the device code

# ============================================
# STEP 6: Publish to MCP Registry
# ============================================
mcp-publisher publish
# Should output: Successfully published

# ============================================
# STEP 7: Verify
# ============================================
Invoke-RestMethod -Uri "https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.xjtlumedia/markdown-formatter" -Proxy "http://127.0.0.1:7890"

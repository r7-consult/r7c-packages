#!/bin/bash

###############################################################################
# Examples Manifest Regeneration Script
#
# This script regenerates examples-manifest.json from the current directory
# structure. Run this whenever you add, remove, or rename example files.
#
# Usage:
#   ./regenerate-manifest.sh          # Run from examples directory
#   bash regenerate-manifest.sh       # Alternative invocation
#
# Output:
#   - Generates ../examples-manifest.json with complete directory structure
#   - Displays statistics (categories, files, total size)
###############################################################################

# Get the directory where this script is located
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Navigate to plugin root (4 levels up from examples/)
# examples/ -> resources/ -> macros_ide/ -> modules/ -> plugin_root/
PLUGIN_ROOT="$(cd "$SCRIPT_DIR/../../../.." && pwd)"

echo "═══════════════════════════════════════════════════════════════════════"
echo "  Examples Manifest Regeneration"
echo "═══════════════════════════════════════════════════════════════════════"
echo ""
echo "Plugin root: $PLUGIN_ROOT"
echo "Examples directory: $SCRIPT_DIR"
echo ""

# Check if Node.js is available
if ! command -v node &> /dev/null; then
    echo "❌ Error: Node.js is not installed or not in PATH"
    echo "   Install Node.js 14+ to use this script"
    exit 1
fi

# Check if generator script exists
GENERATOR_SCRIPT="$PLUGIN_ROOT/modules/macros_ide/scripts/tools/generate-examples-manifest.js"
if [ ! -f "$GENERATOR_SCRIPT" ]; then
    echo "❌ Error: Generator script not found at:"
    echo "   $GENERATOR_SCRIPT"
    exit 1
fi

# Run the generator
echo "🔄 Running manifest generator..."
echo ""
node "$GENERATOR_SCRIPT"
RESULT=$?

echo ""
if [ $RESULT -eq 0 ]; then
    echo "═══════════════════════════════════════════════════════════════════════"
    echo "✅ Manifest regeneration completed successfully!"
    echo "═══════════════════════════════════════════════════════════════════════"
    echo ""
    echo "📁 Output file:"
    echo "   $(cd "$PLUGIN_ROOT/modules/macros_ide/resources" && pwd)/examples-manifest.json"
    echo ""
    echo "💡 Next steps:"
    echo "   1. Restart OnlyOffice to reload the plugin"
    echo "   2. Check that new examples appear in the tree"
    echo "   3. Commit the updated manifest to version control"
else
    echo "═══════════════════════════════════════════════════════════════════════"
    echo "❌ Manifest regeneration failed!"
    echo "═══════════════════════════════════════════════════════════════════════"
    exit 1
fi

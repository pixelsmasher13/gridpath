#!/bin/bash

# Platform-specific build script for GridPath
# Handles resource optimization per platform

set -e  # Exit on error

PLATFORM=${1:-macos}  # Default to macos if not specified

echo "🚀 Building GridPath for $PLATFORM..."

case $PLATFORM in
  macos|darwin)
    echo "📦 Preparing macOS build (excluding Windows resources)..."

    # Temporarily move Windows resources out
    if [ -d "src-tauri/resources/windows" ]; then
      echo "  Moving Windows resources temporarily..."
      mv src-tauri/resources/windows src-tauri/resources-windows-temp
    fi

    # Build for macOS
    echo "  Building..."
    yarn tauri build

    # Restore Windows resources
    if [ -d "src-tauri/resources-windows-temp" ]; then
      echo "  Restoring Windows resources..."
      mv src-tauri/resources-windows-temp src-tauri/resources/windows
    fi

    echo "✅ macOS build complete!"
    ;;

  windows)
    echo "📦 Preparing Windows build (with all resources)..."

    # Windows needs all resources, just build normally
    yarn tauri build

    echo "✅ Windows build complete!"
    ;;

  linux)
    echo "📦 Preparing Linux build (excluding Windows resources)..."

    # Temporarily move Windows resources out
    if [ -d "src-tauri/resources/windows" ]; then
      echo "  Moving Windows resources temporarily..."
      mv src-tauri/resources/windows src-tauri/resources-windows-temp
    fi

    # Build for Linux
    echo "  Building..."
    yarn tauri build

    # Restore Windows resources
    if [ -d "src-tauri/resources-windows-temp" ]; then
      echo "  Restoring Windows resources..."
      mv src-tauri/resources-windows-temp src-tauri/resources/windows
    fi

    echo "✅ Linux build complete!"
    ;;

  all)
    echo "📦 Building for all platforms..."

    # Build macOS first (without Windows resources)
    $0 macos

    # Then build Windows (with all resources)
    $0 windows

    # Then build Linux (without Windows resources)
    $0 linux

    echo "✅ All builds complete!"
    ;;

  *)
    echo "❌ Unknown platform: $PLATFORM"
    echo "Usage: $0 [macos|windows|linux|all]"
    exit 1
    ;;
esac
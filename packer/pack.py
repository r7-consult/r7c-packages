#!/usr/bin/env python
"""
Plugin Packer - Utility for packaging plugins with configurable exclusions
"""

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import time
from datetime import datetime, timezone
from fnmatch import fnmatch
from pathlib import Path
from urllib.parse import urlparse


class PluginPacker:
    """Main class for packing plugins with configurable exclusion patterns"""

    COMMERCIAL_TYPES = {'commercial', 'paid', 'коммерческий'}
    DEFAULT_EXCLUDES = ['deploy/*', 'node_modules/*', '.dev/*']
    MAX_RETRIES = 3
    RETRY_DELAY = 1
    OLD_PATH_PATTERN = r'https://onlyoffice\.github\.io/sdkjs-plugins/v1/[\w\-\./\*]*'
    NEW_PATH = './../v1/'

    def __init__(self, old_mode=False):
        self.old_mode = old_mode
        self.content_dir = "sdkjs-plugins/content/"
        self.artifacts_dir = "artifacts"
        self.integrity_manifest_path = os.path.join("store", "integrity.json")

    def parse_arguments(self):
        """Parse command line arguments"""
        parser = argparse.ArgumentParser(description='Pack plugins')
        parser.add_argument(
            '--old-mode',
            action='store_true',
            help='The old way: put the result in the root of each plugin folder.'
        )
        parser.add_argument(
            '--write-integrity',
            action='store_true',
            help='Generate store/integrity.json for the packed .plugin archives.'
        )
        parser.add_argument(
            '--changed-only',
            action='store_true',
            help='Pack only plugins changed between git refs or missing output archives.'
        )
        parser.add_argument(
            '--base-ref',
            default=None,
            help='Base git ref for --changed-only. Defaults to HEAD~1 when available.'
        )
        parser.add_argument(
            '--head-ref',
            default='HEAD',
            help='Head git ref for --changed-only. Defaults to HEAD.'
        )
        parser.add_argument(
            '--plugin',
            action='append',
            default=[],
            help='Pack only a plugin by name. Can be repeated or comma-separated.'
        )
        parser.add_argument(
            '--verify-integrity',
            action='store_true',
            help='Verify store/integrity.json against generated .plugin archives.'
        )
        return parser.parse_args()

    def get_available_plugins(self):
        """Return available plugin directories keyed by plugin name"""
        if not os.path.exists(self.content_dir):
            return {}

        return {
            plugin_name: os.path.join(self.content_dir, plugin_name)
            for plugin_name in sorted(os.listdir(self.content_dir))
            if os.path.isdir(os.path.join(self.content_dir, plugin_name))
        }

    def parse_plugin_filter(self, plugin_args):
        """Normalize --plugin arguments into a sorted unique list"""
        plugin_names = []
        for plugin_arg in plugin_args or []:
            plugin_names.extend(
                plugin_name.strip()
                for plugin_name in plugin_arg.split(',')
                if plugin_name.strip()
            )
        return sorted(dict.fromkeys(plugin_names))

    def is_commercial_type(self, raw_type):
        """Check whether a type marker means commercial distribution"""
        return str(raw_type or '').strip().lower() in self.COMMERCIAL_TYPES

    def is_valid_external_url(self, value):
        """Check whether a value looks like an external commercial landing URL"""
        if not value or not isinstance(value, str):
            return False
        try:
            parsed = urlparse(value.strip())
        except ValueError:
            return False
        return parsed.scheme in {'http', 'https'} and bool(parsed.netloc)

    def is_commercial_plugin(self, plugin_path):
        """Check whether the plugin should be excluded from .plugin packaging"""
        plugin_manifest = self.read_plugin_manifest(plugin_path)
        variations = plugin_manifest.get("variations", [])
        variation = variations[0] if variations and isinstance(variations[0], dict) else {}
        store = variation.get("store", {}) if isinstance(variation.get("store", {}), dict) else {}

        categories = store.get("categories", [])
        if isinstance(categories, list):
            for category in categories:
                if self.is_commercial_type(category):
                    return True

        if "commercial" in store:
            marker = store.get("commercial")
            if marker is True:
                return True
            if isinstance(marker, str) and marker.strip():
                if self.is_commercial_type(marker) or self.is_valid_external_url(marker):
                    return True
            if isinstance(marker, dict):
                if marker.get("enabled") is True:
                    return True
                for field_name in ("url", "link", "landingUrl", "website"):
                    if self.is_valid_external_url(marker.get(field_name)):
                        return True

        type_value = store.get("type") if isinstance(store.get("type"), str) else plugin_manifest.get("type")
        return self.is_commercial_type(type_value)

    def read_plugin_manifest(self, plugin_path):
        """Read plugin config.json if it exists"""
        config_path = os.path.join(plugin_path, "config.json")
        if not os.path.exists(config_path):
            return {}
        try:
            with open(config_path, 'r', encoding='utf-8') as file:
                return json.load(file)
        except Exception as error:
            print(f"[{os.path.basename(plugin_path)}] Failed to read config.json: {error}")
            return {}

    def get_output_plugin_path(self, plugin_path, plugin_name):
        """Resolve the output .plugin path for the current packing mode"""
        if self.old_mode:
            return os.path.join(plugin_path, "deploy", f"{plugin_name}.plugin")
        return os.path.join(self.artifacts_dir, f"{plugin_name}.plugin")

    def compute_sha256(self, file_path):
        """Compute sha256 checksum for a file"""
        digest = hashlib.sha256()
        with open(file_path, 'rb') as file:
            for chunk in iter(lambda: file.read(1024 * 1024), b''):
                digest.update(chunk)
        return digest.hexdigest()

    def build_integrity_entry(self, plugin_path, plugin_name):
        """Build integrity metadata for a packed plugin"""
        plugin_file_path = self.get_output_plugin_path(plugin_path, plugin_name)
        if not os.path.exists(plugin_file_path):
            return None

        plugin_manifest = self.read_plugin_manifest(plugin_path)
        relative_file_path = os.path.relpath(plugin_file_path, ".").replace("\\", "/")

        return {
            "guid": plugin_manifest.get("guid", ""),
            "version": plugin_manifest.get("version", ""),
            "file": relative_file_path,
            "size": os.path.getsize(plugin_file_path),
            "sha256": self.compute_sha256(plugin_file_path)
        }

    def write_integrity_manifest(self, plugin_entries):
        """Persist integrity metadata for packed plugins"""
        manifest_payload = {
            "generatedAt": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
            "algorithm": "sha256",
            "packages": dict(sorted(plugin_entries.items()))
        }

        os.makedirs(os.path.dirname(self.integrity_manifest_path), exist_ok=True)
        with open(self.integrity_manifest_path, 'w', encoding='utf-8') as file:
            json.dump(manifest_payload, file, ensure_ascii=False, indent=2)
            file.write("\n")

        print(f"Integrity manifest written: {self.integrity_manifest_path}")

    def delete_output_plugin(self, plugin_path, plugin_name):
        """Delete a generated .plugin archive for the current mode if it exists"""
        plugin_file_path = self.get_output_plugin_path(plugin_path, plugin_name)
        if os.path.exists(plugin_file_path):
            os.remove(plugin_file_path)
            print(f"[SKIP] Removed commercial archive: {plugin_name}")

    def cleanup_commercial_output_plugins(self):
        """Remove generated archives for commercial plugins in the current mode"""
        commercial_plugins = []
        for plugin_name, plugin_path in self.get_available_plugins().items():
            if self.is_commercial_plugin(plugin_path):
                commercial_plugins.append(plugin_name)
                self.delete_output_plugin(plugin_path, plugin_name)
        return commercial_plugins

    def read_integrity_manifest(self):
        """Read and normalize the integrity manifest from disk"""
        if not os.path.exists(self.integrity_manifest_path):
            raise FileNotFoundError(f"Integrity manifest is missing: {self.integrity_manifest_path}")
        with open(self.integrity_manifest_path, 'r', encoding='utf-8') as file:
            payload = json.load(file)
        if payload.get("algorithm") != "sha256":
            raise ValueError("Integrity manifest must use sha256")
        packages = payload.get("packages")
        if not isinstance(packages, dict):
            raise ValueError("Integrity manifest packages must be an object")
        return payload

    def verify_integrity_manifest(self):
        """Verify manifest contents against actual .plugin archives"""
        manifest = self.read_integrity_manifest()
        available_plugins = self.get_available_plugins()
        packages = manifest.get("packages", {})
        errors = []

        for plugin_name, plugin_path in available_plugins.items():
            is_commercial = self.is_commercial_plugin(plugin_path)
            entry = packages.get(plugin_name)
            plugin_file_path = self.get_output_plugin_path(plugin_path, plugin_name)
            plugin_manifest = self.read_plugin_manifest(plugin_path)

            if is_commercial:
                if entry:
                    errors.append(f"{plugin_name}: commercial plugin must not be present in integrity manifest")
                if os.path.exists(plugin_file_path):
                    errors.append(f"{plugin_name}: commercial plugin must not have generated archive {plugin_file_path}")
                continue

            if not os.path.exists(plugin_file_path):
                if entry:
                    errors.append(f"{plugin_name}: manifest entry exists but archive is missing")
                continue

            if not entry:
                errors.append(f"{plugin_name}: archive exists but manifest entry is missing")
                continue

            expected_file = os.path.relpath(plugin_file_path, ".").replace("\\", "/")
            actual_size = os.path.getsize(plugin_file_path)
            actual_sha256 = self.compute_sha256(plugin_file_path)
            expected_guid = plugin_manifest.get("guid", "")
            expected_version = plugin_manifest.get("version", "")

            if entry.get("file") != expected_file:
                errors.append(f"{plugin_name}: manifest file path mismatch")
            if int(entry.get("size", 0)) != actual_size:
                errors.append(f"{plugin_name}: manifest size mismatch")
            if str(entry.get("sha256", "")).lower() != actual_sha256:
                errors.append(f"{plugin_name}: manifest sha256 mismatch")
            if entry.get("guid", "") != expected_guid:
                errors.append(f"{plugin_name}: manifest guid mismatch")
            if entry.get("version", "") != expected_version:
                errors.append(f"{plugin_name}: manifest version mismatch")

        for plugin_name in packages.keys():
            if plugin_name not in available_plugins:
                errors.append(f"{plugin_name}: manifest entry points to unknown plugin")

        if errors:
            for error in errors:
                print(f"[FAIL] {error}")
            return False

        print(f"[OK] Integrity manifest verified: {len(packages)} package(s)")
        return True

    def collect_integrity_entries(self):
        """Build integrity metadata from existing .plugin archives"""
        plugin_entries = {}
        for plugin_name, plugin_path in self.get_available_plugins().items():
            if self.is_commercial_plugin(plugin_path):
                continue
            entry = self.build_integrity_entry(plugin_path, plugin_name)
            if entry:
                plugin_entries[plugin_name] = entry
        return plugin_entries

    def get_missing_output_plugins(self):
        """Return plugins that do not have a .plugin archive for the current mode"""
        missing_plugins = []
        for plugin_name, plugin_path in self.get_available_plugins().items():
            if self.is_commercial_plugin(plugin_path):
                continue
            plugin_file_path = self.get_output_plugin_path(plugin_path, plugin_name)
            if not os.path.exists(plugin_file_path):
                missing_plugins.append(plugin_name)
        return missing_plugins

    def is_zero_ref(self, ref):
        """Return True for GitHub's zero SHA used when there is no before commit"""
        return bool(ref) and set(ref) == {'0'}

    def git_ref_exists(self, ref):
        """Check if a git ref exists locally"""
        if not ref or self.is_zero_ref(ref):
            return False

        result = subprocess.run(
            ["git", "rev-parse", "--verify", f"{ref}^{{commit}}"],
            capture_output=True,
            text=True
        )
        return result.returncode == 0

    def resolve_base_ref(self, base_ref, head_ref):
        """Resolve the base ref for changed-only packing"""
        if self.git_ref_exists(base_ref):
            return base_ref

        fallback_base = f"{head_ref or 'HEAD'}~1"
        if self.git_ref_exists(fallback_base):
            print(f"Base ref is not available, using {fallback_base}")
            return fallback_base

        print("Base ref is not available; packing all plugins")
        return None

    def get_changed_files(self, base_ref, head_ref):
        """Return changed files between git refs, or None when diff cannot be resolved"""
        resolved_base_ref = self.resolve_base_ref(base_ref, head_ref)
        if not resolved_base_ref:
            return None

        result = subprocess.run(
            [
                "git",
                "diff",
                "--name-only",
                "--diff-filter=ACMRT",
                resolved_base_ref,
                head_ref or "HEAD",
                "--",
                self.content_dir,
                "store/config.json",
                "packer/pack.py"
            ],
            capture_output=True,
            text=True
        )

        if result.returncode != 0:
            print(f"Failed to get changed files: {result.stderr.strip()}")
            return None

        return [
            changed_file.strip()
            for changed_file in result.stdout.splitlines()
            if changed_file.strip()
        ]

    def get_changed_plugins(self, base_ref, head_ref):
        """Resolve changed plugin names, or None when all plugins must be packed"""
        changed_files = self.get_changed_files(base_ref, head_ref)
        if changed_files is None:
            return None

        content_prefix = self.content_dir.replace("\\", "/").rstrip("/") + "/"
        changed_plugins = set()

        for changed_file in changed_files:
            normalized_path = changed_file.replace("\\", "/")

            if normalized_path == "store/config.json" or normalized_path.startswith("packer/"):
                print(f"Global packaging input changed: {normalized_path}; packing all plugins")
                return None

            if not normalized_path.startswith(content_prefix):
                continue

            relative_plugin_path = normalized_path[len(content_prefix):]
            parts = relative_plugin_path.split("/")
            if len(parts) < 2:
                continue

            plugin_name = parts[0]
            if parts[1] == "deploy":
                continue

            changed_plugins.add(plugin_name)

        return sorted(changed_plugins)

    def replace_html_paths(self, file_path):
        """Replace paths in HTML files before packing"""
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()

            new_content = re.sub(
                self.OLD_PATH_PATTERN,
                lambda match: self.NEW_PATH + match.group(0).split('/v1/')[1],
                content
            )

            if new_content != content:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(new_content)
                return True
            return False

        except Exception as error:
            print(f"  Error processing HTML file {file_path}: {error}")
            return False

    def process_html_files(self, directory):
        """Find and process all HTML files in directory"""
        html_files_processed = 0
        for root, dirs, files in os.walk(directory):
            for file_name in files:
                if file_name.lower().endswith('.html'):
                    file_path = os.path.join(root, file_name)
                    if self.replace_html_paths(file_path):
                        html_files_processed += 1
        return html_files_processed

    def safe_rename(self, src, dst, max_retries=MAX_RETRIES, delay=RETRY_DELAY):
        """Safely rename file with retry logic for locked files"""
        for attempt in range(max_retries):
            try:
                Path(src).rename(dst)
                return True
            except PermissionError:
                if attempt < max_retries - 1:
                    print(f"  File busy, retrying in {delay} second(s)...")
                    time.sleep(delay)
                else:
                    print(f"  Failed to rename after {max_retries} attempts")
                    return False
        return False

    def delete_dir(self, path, max_retries=MAX_RETRIES, delay=RETRY_DELAY):
        """Safely delete directory with retry logic"""
        if not os.path.exists(path):
            return True

        for attempt in range(max_retries):
            try:
                shutil.rmtree(path)
                return True
            except PermissionError:
                if attempt < max_retries - 1:
                    print(f"Attempt {attempt + 1}: File in use... {path}")
                    time.sleep(delay)
                else:
                    print(f"Failed to delete {path}")
                    return False
        return False

    def get_plugin_excludes(self, plugin_path):
        """Get exclusion patterns for plugin from config.json"""
        plugin_config_path = os.path.join(plugin_path, ".dev", "config.json")

        excludes_set = set(self.DEFAULT_EXCLUDES)

        if os.path.exists(plugin_config_path):
            try:
                with open(plugin_config_path, 'r', encoding='utf-8') as file:
                    config = json.load(file)
                    custom_excludes = config.get("excludes", [])
                    excludes_set.update(custom_excludes)
            except (json.JSONDecodeError, Exception) as error:
                print(f"[{os.path.basename(plugin_path)}] Error reading config: {error}")

        return list(excludes_set)

    def should_exclude(self, path, base_dir, excludes):
        """Check if path should be excluded based on patterns"""
        if not excludes:
            return False

        if os.path.isabs(path):
            relative_path = os.path.relpath(path, base_dir)
        else:
            relative_path = path

        relative_path = relative_path.replace('\\', '/')

        for pattern in excludes:
            pattern = pattern.replace('\\', '/')
            if (
                fnmatch(relative_path, pattern) or
                fnmatch(relative_path + '/', pattern + '/')
            ):
                return True

        return False

    def copy_filtered_files(self, source_dir, dest_dir, excludes):
        """Copy files with exclusion patterns applied"""
        for root, dirs, files in os.walk(source_dir):
            dirs[:] = [
                dir_name for dir_name in dirs
                if not self.should_exclude(os.path.join(root, dir_name), source_dir, excludes)
            ]

            for file_name in files:
                source_file = os.path.join(root, file_name)
                relative_path = os.path.relpath(source_file, source_dir)

                if not self.should_exclude(relative_path, source_dir, excludes):
                    dest_file = os.path.join(dest_dir, relative_path)
                    os.makedirs(os.path.dirname(dest_file), exist_ok=True)
                    shutil.copy2(source_file, dest_file)

    def create_plugin_archive(self, source_dir, plugin_name, output_dir, excludes):
        """Create .plugin archive from source directory"""
        temp_dir = os.path.join(output_dir, f"temp_{plugin_name}")
        os.makedirs(temp_dir, exist_ok=True)

        try:
            self.copy_filtered_files(source_dir, temp_dir, excludes or [])

            if not any(os.scandir(temp_dir)):
                print(f"[{plugin_name}] No files to pack after filtering")
                return False

            self.process_html_files(temp_dir)

            zip_file = os.path.join(output_dir, plugin_name)
            zip_path = shutil.make_archive(zip_file, 'zip', temp_dir)
            plugin_file_path = Path(zip_path).with_suffix(".plugin")

            if self.safe_rename(zip_path, plugin_file_path):
                print(f"[OK] Created: {plugin_name}")
                return True

            print(f"[FAIL] Failed to create: {plugin_name}")
            return False

        except Exception as error:
            print(f"[{plugin_name}] Error: {error}")
            return False
        finally:
            self.delete_dir(temp_dir)

    def pack_plugin_new_mode(self, plugin_path, plugin_name):
        """Pack plugin in new mode (artifacts directory)"""
        excludes = self.get_plugin_excludes(plugin_path)
        return self.create_plugin_archive(plugin_path, plugin_name, self.artifacts_dir, excludes)

    def pack_plugin_old_mode(self, plugin_path, plugin_name):
        """Pack plugin in old mode (plugin deploy directory)"""
        excludes = self.get_plugin_excludes(plugin_path)
        destination_path = os.path.join(plugin_path, "deploy")

        if os.path.exists(destination_path):
            self.delete_dir(destination_path)

        return self.create_plugin_archive(plugin_path, plugin_name, destination_path, excludes)

    def pack_plugins(self, plugin_names=None):
        """Main packing method"""
        if not os.path.exists(self.content_dir):
            print(f"Content directory {self.content_dir} does not exist")
            return []

        os.makedirs(self.artifacts_dir, exist_ok=True)
        available_plugins = self.get_available_plugins()
        packed_plugins = []

        selected_plugin_names = (
            sorted(available_plugins.keys())
            if plugin_names is None
            else sorted(dict.fromkeys(plugin_names))
        )

        if not selected_plugin_names:
            print("No plugins selected for packing")
            return packed_plugins

        for plugin_name in selected_plugin_names:
            plugin_path = available_plugins.get(plugin_name)
            if not plugin_path:
                print(f"[WARN] Skipping unknown plugin: {plugin_name}")
                continue
            if self.is_commercial_plugin(plugin_path):
                print(f"[SKIP] Commercial plugin: {plugin_name}")
                continue

            if self.old_mode:
                packed = self.pack_plugin_old_mode(plugin_path, plugin_name)
            else:
                packed = self.pack_plugin_new_mode(plugin_path, plugin_name)

            if packed:
                packed_plugins.append(plugin_name)

        return packed_plugins

    def run(self):
        """Run the plugin packer"""
        args = self.parse_arguments()
        self.old_mode = args.old_mode
        selected_plugins = self.parse_plugin_filter(args.plugin) or None
        verify_only = args.verify_integrity and not args.write_integrity and not args.changed_only and not selected_plugins

        if not verify_only:
            self.cleanup_commercial_output_plugins()

        if args.changed_only:
            changed_plugins = self.get_changed_plugins(args.base_ref, args.head_ref)
            missing_plugins = self.get_missing_output_plugins()
            if changed_plugins is not None:
                plugins_to_pack = sorted(set(changed_plugins) | set(missing_plugins))
                if missing_plugins:
                    print(f"Missing plugin archives: {', '.join(missing_plugins)}")
                selected_plugins = plugins_to_pack if selected_plugins is None else [
                    plugin_name
                    for plugin_name in selected_plugins
                    if plugin_name in plugins_to_pack
                ]

        packed_plugins = []
        if not verify_only:
            packed_plugins = self.pack_plugins(selected_plugins)
            print(f"Packed plugins: {len(packed_plugins)}")

        if args.write_integrity:
            plugin_entries = self.collect_integrity_entries()
            self.write_integrity_manifest(plugin_entries)

        if args.verify_integrity:
            if not self.verify_integrity_manifest():
                raise SystemExit(1)


def main():
    """Main entry point"""
    packer = PluginPacker()
    packer.run()


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
Security utilities for PowerPoint template processing.

Provides path validation, file size checking, and secure exception handling
to prevent common vulnerabilities like path traversal and zip bombs.
"""

import os
import sys
from pathlib import Path
from typing import Optional


class SecurityError(Exception):
    """Raised when a security check fails."""
    pass


class PathTraversalError(SecurityError):
    """Raised when path traversal attempt is detected."""
    pass


class FileSizeError(SecurityError):
    """Raised when file size exceeds limits."""
    pass


# Security configuration
MAX_FILE_SIZE_MB = 100  # Maximum allowed file size in MB
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024


def validate_file_path(file_path: str, must_exist: bool = True,
                       allowed_extensions: Optional[list] = None,
                       base_dir: Optional[str] = None) -> Path:
    """
    Validate and sanitize a file path to prevent security vulnerabilities.

    Args:
        file_path: Path to validate
        must_exist: Whether the file must exist
        allowed_extensions: List of allowed file extensions (e.g., ['.pptx', '.json'])
        base_dir: Optional base directory to restrict operations to

    Returns:
        Validated Path object

    Raises:
        PathTraversalError: If path traversal attempt detected
        FileNotFoundError: If must_exist=True and file doesn't exist
        ValueError: If file extension not allowed
    """
    try:
        # Convert to Path object and resolve to absolute path
        path = Path(file_path).resolve()
    except (OSError, RuntimeError) as e:
        raise PathTraversalError(f"Invalid path: {e}")

    # Check if path exists (if required)
    if must_exist and not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # Prevent path traversal by checking for parent directory references
    # The resolve() call above handles symlinks and normalizes the path
    try:
        # Check if the resolved path contains suspicious patterns
        path_str = str(path)
        if ".." in Path(file_path).parts:
            raise PathTraversalError(
                f"Path traversal attempt detected: {file_path}"
            )
    except ValueError as e:
        raise PathTraversalError(f"Invalid path structure: {e}")

    # If base_dir specified, ensure path is within it
    if base_dir:
        try:
            base = Path(base_dir).resolve()
            # Check if path is relative to base_dir
            path.relative_to(base)
        except (ValueError, OSError):
            raise PathTraversalError(
                f"Path outside allowed directory. Path: {path}, Base: {base_dir}"
            )

    # Validate file extension if specified
    if allowed_extensions and path.suffix.lower() not in allowed_extensions:
        raise ValueError(
            f"Invalid file extension: {path.suffix}. "
            f"Allowed: {', '.join(allowed_extensions)}"
        )

    return path


def check_file_size(file_path: Path, max_size_bytes: int = MAX_FILE_SIZE_BYTES) -> None:
    """
    Check if file size is within acceptable limits.

    Prevents zip bomb attacks and excessive resource consumption.

    Args:
        file_path: Path to file to check
        max_size_bytes: Maximum allowed file size in bytes

    Raises:
        FileSizeError: If file exceeds size limit
    """
    if not file_path.exists():
        return  # Will be caught by validate_file_path

    file_size = file_path.stat().st_size

    if file_size > max_size_bytes:
        max_size_mb = max_size_bytes / (1024 * 1024)
        actual_size_mb = file_size / (1024 * 1024)
        raise FileSizeError(
            f"File too large: {actual_size_mb:.2f}MB exceeds "
            f"limit of {max_size_mb:.2f}MB"
        )


def validate_input_file(file_path: str, allowed_extensions: list,
                        max_size_bytes: int = MAX_FILE_SIZE_BYTES) -> Path:
    """
    Comprehensive validation for input files.

    Combines path validation and size checking.

    Args:
        file_path: Path to validate
        allowed_extensions: List of allowed extensions
        max_size_bytes: Maximum file size

    Returns:
        Validated Path object

    Raises:
        SecurityError: If validation fails
    """
    # Validate path
    path = validate_file_path(
        file_path,
        must_exist=True,
        allowed_extensions=allowed_extensions
    )

    # Check file size
    check_file_size(path, max_size_bytes)

    return path


def validate_output_file(file_path: str, allowed_extensions: list) -> Path:
    """
    Validate output file path.

    Ensures output path is safe and uses allowed extension.

    Args:
        file_path: Path to validate
        allowed_extensions: List of allowed extensions

    Returns:
        Validated Path object

    Raises:
        SecurityError: If validation fails
    """
    # Validate path (output file doesn't need to exist)
    path = validate_file_path(
        file_path,
        must_exist=False,
        allowed_extensions=allowed_extensions
    )

    # Ensure parent directory exists or can be created
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
    except (OSError, PermissionError) as e:
        raise SecurityError(f"Cannot create output directory: {e}")

    return path


def safe_file_read(file_path: Path) -> bytes:
    """
    Safely read file contents with size validation.

    Args:
        file_path: Path to read

    Returns:
        File contents as bytes

    Raises:
        FileSizeError: If file is too large
        OSError: If file cannot be read
    """
    # Check size before reading
    check_file_size(file_path)

    try:
        return file_path.read_bytes()
    except PermissionError as e:
        raise SecurityError(f"Permission denied reading file: {e}")
    except OSError as e:
        raise SecurityError(f"Error reading file: {e}")

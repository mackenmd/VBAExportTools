# Lesson:Python
# Python modules commonly group standard-library imports first, then third-party
# imports. Unlike Node's package names, `from package import name` binds a
# particular exported name directly in this module's namespace. `pathlib.Path`
# is an object-oriented path API; it complements rather than replaces the older
# string-based `os.path` functions used elsewhere in this established script.
import os
import shutil
import sys
import subprocess
import logging
import tempfile
from pathlib import Path
from datetime import datetime
from oletools.olevba import VBA_Parser


# Lesson:Python
# A class can inherit from a framework class just as in JavaScript. `__init__`
# is Python's constructor hook, and `super()` calls the parent implementation.
# Python methods declare `self` explicitly; it plays the role JavaScript supplies
# implicitly as `this`. This handler records errors because the library reports
# some extraction failures through logging instead of exceptions.
class _ExtractionErrors(logging.Handler):
    """Remember errors that oletools logs instead of raising."""

    def __init__(self):
        super().__init__(logging.ERROR)
        self.errors = []

    def emit(self, record):
        self.errors.append(record.getMessage())


# Lesson:Python
# `def` introduces a function, and its parameters are ordinary local bindings.
# Python has no braces: indentation defines the blocks belonging to `try`, `if`,
# and loops. `None` is Python's single null-like object; use `is None` rather
# than equality when testing for it, since identity is the intended question.
def _stage_and_replace(source_path, destination_dir):
    destination_dir = os.path.abspath(destination_dir)
    staging = None
    backup = None
    try:
        # Lesson:Python
        # `Path(...)` wraps a path in a path object, and `/` has been overloaded
        # to join path components. Here `is_relative_to` makes a containment
        # decision after `resolve`; this is more precise than a JavaScript string
        # prefix test, which can confuse similarly named directories.
        if Path(tempfile.gettempdir()).resolve().is_relative_to(Path(destination_dir).resolve()):
            raise ValueError("Temporary directories must be outside the destination directory.")
        staging = tempfile.mkdtemp(prefix="vba-stage-")

        errors = _ExtractionErrors()
        logger = logging.getLogger("oletools")
        logger.addHandler(errors)
        parser = None
        # Lesson:Python
        # `try`/`finally` is the direct analogue of JavaScript `try`/`finally`:
        # the `finally` block runs whether the protected work returns or raises.
        # The nested structure ensures both the parser and the temporary logging
        # handler are released, even if construction or extraction fails.
        try:
            parser = VBA_Parser(source_path)
            modules = list(parser.extract_macros()) if parser.detect_vba_macros() else []
        finally:
            try:
                if parser is not None:
                    parser.close()
            finally:
                logger.removeHandler(errors)
        if errors.errors:
            raise ValueError("VBA extraction failed: " + "; ".join(errors.errors))
        if not modules:
            return []

        exported_files = []
        seen = set()
        for _, _, filename, code in sorted(modules, key=lambda item: item[2].lower()):
            # Lesson:Python
            # Tuple unpacking binds each element of a returned record in one
            # operation; `_` conventionally names values deliberately ignored.
            # `lambda` is a small anonymous function, much like a JavaScript
            # arrow callback. `set` gives efficient membership checks, while the
            # starred generator expressions extend the reserved-name set without
            # building temporary lists.
            # Never interpret parser output as a path, device, or NTFS stream.
            if (not filename or any(c in filename for c in '<>:"/\\|?*')
                    or any(ord(c) < 32 for c in filename)
                    or filename.endswith((" ", "."))
                    or filename.split(".")[0].upper() in
                    {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}):
                raise ValueError(f"Unsafe VBA module filename: {filename!r}")
            if filename.casefold() in seen:
                raise ValueError(f"Duplicate VBA module filename: {filename}")
            seen.add(filename.casefold())
            if isinstance(code, str):
                code = code.encode("utf-8")
            # Lesson:Python
            # `with` uses a context manager: entering the block acquires a
            # resource and leaving it reliably closes it, including on errors.
            # It is Python's idiomatic spelling of the resource-management
            # pattern often written with `try/finally` in Node. Mode `"xb"`
            # means binary, exclusive creation, so an unexpected existing file
            # is an error rather than something silently overwritten.
            with open(os.path.join(staging, filename), "xb") as output:
                output.write(code)
            exported_files.append(filename)

        existing = []
        if os.path.isdir(destination_dir):
            with os.scandir(destination_dir) as entries:
                for entry in entries:
                    if entry.is_dir():
                        if entry.name.casefold() in seen:
                            raise ValueError(f"Module conflicts with a destination subdirectory: {entry.name}")
                        continue
                    if entry.is_symlink() or not entry.is_file(follow_symlinks=False):
                        raise ValueError(f"Refusing to replace a linked or special destination file: {entry.name}")
                    existing.append(entry.name)

        for filename in set(existing + exported_files):
            target = os.path.join(destination_dir, filename)
            if (os.path.normcase(os.path.realpath(target)) == os.path.normcase(os.path.realpath(source_path))
                    or (os.path.exists(target) and os.path.samefile(source_path, target))):
                raise ValueError(f"Source file conflicts with an exporter-owned destination file: {target}")

        backup = tempfile.mkdtemp(prefix="vba-backup-")
        for filename in existing:
            shutil.copy2(os.path.join(destination_dir, filename), os.path.join(backup, filename))

        installed = []
        removed = []
        try:
            os.makedirs(destination_dir, exist_ok=True)
            for filename in existing:
                os.unlink(os.path.join(destination_dir, filename))
                removed.append(filename)
            for filename in exported_files:
                # Exclusive creation prevents overwriting a newly appeared file.
                with open(os.path.join(destination_dir, filename), "xb") as output:
                    installed.append(filename)
                    with open(os.path.join(staging, filename), "rb") as staged:
                        shutil.copyfileobj(staged, output)
        except Exception:
            # Lesson:Python
            # `except Exception` catches ordinary application failures but not
            # control-flow exceptions such as `KeyboardInterrupt` and
            # `SystemExit`. The bare `raise` below re-raises the same exception
            # after rollback, preserving its traceback—an idiomatic Python
            # alternative to saving and throwing a caught JavaScript error.
            recovery_errors = []
            for filename in installed:
                try:
                    os.unlink(os.path.join(destination_dir, filename))
                except OSError as error:
                    recovery_errors.append(str(error))
            for filename in removed:
                try:
                    target = os.path.join(destination_dir, filename)
                    with open(target, "xb") as output:
                        with open(os.path.join(backup, filename), "rb") as saved:
                            shutil.copyfileobj(saved, output)
                    shutil.copystat(os.path.join(backup, filename), target)
                except OSError as error:
                    recovery_errors.append(str(error))
            if recovery_errors:
                print(f"⚠️ Recovery incomplete. Backup retained at {backup}: {'; '.join(recovery_errors)}")
                backup = None  # Preserve recovery material for manual restoration.
            raise
        return exported_files
    finally:
        # This final cleanup is intentionally separate from the rollback above:
        # it applies to success, an early return, and an exception alike.
        for directory in (staging, backup):
            if directory is not None:
                try:
                    shutil.rmtree(directory)
                except OSError as error:
                    print(f"⚠️ Could not remove temporary directory {directory}: {error}")


def export_vba_code_internal():
    """
    Exports VBA code from a supported VBA file using the Python library,
    sorts modules alphabetically, summarizes export count,
    and optionally commits to Git.
    """
    if len(sys.argv) < 3:
        print("❌ Error: Missing source VBA file path or destination export folder.")
        print("Usage: export_vba_code.py <source file> <destination folder> [commit message|nocommit]")
        return

    source_path = sys.argv[1]
    destination_dir = sys.argv[2]

    if not os.path.isfile(source_path):
        print(f"❌ Error: Source VBA file not found at {source_path}")
        return

    if os.path.splitext(source_path)[1].lower() not in {".otm", ".xlsm", ".xlam"}:
        print("❌ Error: Source VBA file must have a .otm, .xlsm, or .xlam extension.")
        return

    if os.path.exists(destination_dir) and not os.path.isdir(destination_dir):
        print(f"❌ Error: Destination path exists but is not a directory: {destination_dir}")
        return

    # Lesson:Python
    # A conditional expression has the form `value_if_true if condition else
    # value_if_false`, the reverse order of JavaScript's ternary operator.
    # `" ".join(...)` is an idiomatic way to combine a sequence of strings.
    # The datetime format specification after `:` inside an f-string delegates
    # to the object's formatting protocol rather than being JavaScript syntax.
    # Determine commit message or NOCOMMIT flag from remaining command-line args
    if len(sys.argv) > 3:
        commit_msg = " ".join(sys.argv[3:])
    else:
        commit_msg = f"Auto-export VBA {datetime.now():%Y-%m-%d %H:%M}"

    no_commit = commit_msg.strip().lower() == "nocommit"

    try:
        exported_files = _stage_and_replace(source_path, destination_dir)
        if exported_files:
            print(f"\n✅ Successfully exported {len(exported_files)} VBA modules to:")
            print(f"   {destination_dir}\n")

            for name in exported_files:
                print(f"   - {name}")

            # Optionally commit to Git
            if not no_commit:
                print(f"\n📘 Committing to git with message: {commit_msg}")
                subprocess.run(["git", "-C", destination_dir, "add", "."], check=False)
                subprocess.run(["git", "-C", destination_dir, "commit", "-m", commit_msg], check=False)
                print("✅ Git commit complete.\n")
            else:
                print("🚫 NOCOMMIT flag detected — skipping git commit.\n")

        else:
            print("⚠️ No VBA macros found in the source file.")

    except Exception as e:
        # Lesson:Python
        # `as e` binds the caught exception for this handler. Converting it in an
        # f-string calls its human-facing string representation. At this command
        # boundary the exporter reports the problem instead of re-raising it.
        print(f"❌ An error occurred during VBA export: {e}")


# Lesson:Python
# Every Python file has a module name. When this file is run directly, Python
# assigns `"__main__"`; when it is imported, it assigns the import name instead.
# This guard is the idiomatic equivalent of a Node entry-point check and keeps
# command-line work from running merely because another module imported it.
if __name__ == "__main__":
    export_vba_code_internal()

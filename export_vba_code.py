import os
import shutil
import sys
import subprocess
from datetime import datetime
from oletools.olevba import VBA_Parser


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

    # Determine commit message or NOCOMMIT flag from remaining command-line args
    if len(sys.argv) > 3:
        commit_msg = " ".join(sys.argv[3:])
    else:
        commit_msg = f"Auto-export VBA {datetime.now():%Y-%m-%d %H:%M}"

    no_commit = commit_msg.strip().lower() == "nocommit"

    # Ensure the destination directory is clean and exists
    if os.path.exists(destination_dir):
        for filename in os.listdir(destination_dir):
            file_path = os.path.join(destination_dir, filename)
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
    else:
        os.makedirs(destination_dir, exist_ok=True)

    # Use VBA_Parser to extract macros
    try:
        vbaparser = VBA_Parser(source_path)
        exported_files = []

        if vbaparser.detect_vba_macros():
            # Sort modules alphabetically by filename
            for _, stream_path, vba_filename, vba_code in sorted(vbaparser.extract_macros(), key=lambda x: x[2].lower()):
                file_path = os.path.join(destination_dir, vba_filename)

                # Ensure vba_code is in bytes format before writing
                if isinstance(vba_code, str):
                    vba_code = vba_code.encode("utf-8")

                try:
                    with open(file_path, "wb") as f:
                        f.write(vba_code)
                    exported_files.append(vba_filename)
                except IOError as e:
                    print(f"⚠️ Error writing file {file_path}: {e}")

            vbaparser.close()
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
            print("⚠️ No VBA macros found in the OTM file.")
            vbaparser.close()

    except Exception as e:
        print(f"❌ An error occurred during VBA extraction: {e}")


if __name__ == "__main__":
    export_vba_code_internal()

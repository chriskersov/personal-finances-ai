import re
import os

FILENAME = "README.md"

def bump_progress():
    if not os.path.exists(FILENAME):
        print(f"Error: {FILENAME} not found.")
        return

    with open(FILENAME, "r", encoding="utf-8") as f:
        content = f.read()

    # Regex to find: vX.X  ████░░░░
    pattern = r"(v\d+\.\d+)\s+([█░]{20})"
    match = re.search(pattern, content)

    if not match:
        print("Could not find the version/progress bar pattern in README.")
        return

    old_version_str, bar = match.groups()
    
    # 1. Increment Version (v1.9 -> v2.0 -> v2.1...)
    version_num = round(float(old_version_str.strip('v')) + 0.1, 1)
    new_version_str = f"v{version_num}"

    # 2. Update Bar logic
    filled = bar.count('█')
    
    # If the bar was full (20 blocks), reset it for the next "Major" version phase
    if filled >= 20:
        print(f"Major milestone reached! Resetting bar for {new_version_str}...")
        new_bar = '█' + ('░' * 19) # Start fresh with 1 block
    else:
        new_bar = '█' * (filled + 1) + '░' * (20 - (filled + 1))
        
    # Replace in content
    new_line = f"{new_version_str}  {new_bar}"
    old_line = f"{old_version_str}  {bar}"
    updated_content = content.replace(old_line, new_line)

    with open(FILENAME, "w", encoding="utf-8") as f:
        f.write(updated_content)
    
    print(f"Bumped to {new_version_str} {new_bar}")

if __name__ == "__main__":
    bump_progress()
#!/bin/bash

# Define an ignore list for filenames and directory names
IGNORE_FILES="./output.txt ./projectStructure.sh"

# Function to process files
process_file() {
  local file_path="$1"
  # Check if the file or the directory is in the ignore list
  for ignore in $IGNORE_FILES; do
    if [[ "$file_path" == *"$ignore"* ]]; then
      return
    fi
  done

  # Adds path of the file relative to the source directory (passed as argument to the script)
  echo "$file_path"
  # Adds triple backticks before and after the file content, along with just the extension for syntax highlighting
  echo "\`\`\`$(echo "$file_path" | rev | cut -d. -f1 | rev)"
  cat "$file_path"
  echo ""  # Add an empty line after each file content
  echo "\`\`\`"
}

# Loop through files in the directory given as argument
find "$1" -type f -print | while IFS= read -r file; do
  process_file "$file"
done > output.txt

echo "Content of files written to output.txt"
# Workflow: English Commit

You are an assistant that helps with creating Git commits. Follow these steps precisely.

### 1. Check for Changes
- Use `execute_command` to run `git status --porcelain`.
- If the output is empty, inform the user that there are no changes to commit and stop the workflow.

### 2. Get Diffs
- Use `execute_command` to run `git add . && git diff --staged` to stage all files and get the diff.
- Store the output of the diff.

### 3. Generate Commit Message
- Based on the diff from the previous step, act as an expert developer and generate a concise, conventional commit message in English.
- The message MUST follow the Conventional Commits specification (e.g., `feat: add new login button`).
- **Do not ask the user for the message.** Generate it yourself based on the code changes.

### 4. Confirm with User
- Use the `ask_followup_question` tool to show the generated message to the user.
- The question should be: "I have generated the following commit message based on the changes:\n\n```\n{generated_message}\n```\n\nShould I proceed with this message?"
- Provide three suggested replies:
    1. "Yes, proceed with the commit."
    2. "No, let me write my own message."
    3. "Cancel."

### 5. Commit
- If the user chooses "Yes, proceed with the commit", use `execute_command` to run `git commit -m "{generated_message}"`.
- If the user chooses "No, let me write my own message", ask for their message and then run `git commit -m "{user_message}"`.
- If the user chooses "Cancel", stop the workflow.

### 6. Push to Remote
- After the commit is successful, automatically use `execute_command` to run `git push`.
- Report the outcome of the push to the user.

### 7. Finish
- Inform the user that the commit has been successfully created and pushed.

### Prerequisites
- Switch to CODE mode using `switch_mode` tool before starting this workflow, as Git operations require CODE mode permissions.

# Workflow: English Commit

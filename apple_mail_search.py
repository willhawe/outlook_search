import subprocess
import sys

def search_apple_mail(search_string):
    """
    Search Apple Mail using AppleScript
    Note: This works with Apple Mail, which can be connected to Outlook accounts
    """
    apple_script = f'''
    tell application "Mail"
        set searchResults to {{}}
        set allMessages to messages of inbox
        
        repeat with currentMessage in allMessages
            set messageContent to content of currentMessage
            set messageSubject to subject of currentMessage
            
            if messageContent contains "{search_string}" or messageSubject contains "{search_string}" then
                set messageInfo to {{subject:messageSubject, sender:sender of currentMessage, date:date received of currentMessage}}
                copy messageInfo to end of searchResults
            end if
        end repeat
        
        return searchResults
    end tell
    '''
    
    try:
        result = subprocess.run(['osascript', '-e', apple_script], 
                              capture_output=True, text=True, check=True)
        
        if result.stdout.strip():
            print(f"Found emails containing '{search_string}':")
            print(result.stdout)
        else:
            print(f"No emails found containing '{search_string}'")
            
    except subprocess.CalledProcessError as e:
        print(f"Error running AppleScript: {e}")
        print("Make sure Apple Mail is installed and running")

def main():
    print("Apple Mail Search Tool")
    print("=" * 25)
    print("Note: This searches Apple Mail, which can be connected to Outlook accounts")
    
    search_term = input("Enter search term: ").strip()
    
    if not search_term:
        print("Please enter a search term.")
        return
    
    search_apple_mail(search_term)

if __name__ == "__main__":
    main() 
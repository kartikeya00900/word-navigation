import win32com.client as win32
import speech_recognition as sr

listener=sr.Recognizer()



def take_command():
    try:
        with sr.Microphone() as source:
            print('listening...')
            voice = listener.listen(source)
            command = listener.recognize_google(voice)
            command = command.lower()
        
    except:
        pass
    return command

# Open Microsoft Word
word = win32.Dispatch("Word.Application")
word.Visible = True

# Create a new document
doc = word.Documents.Add()

# Ask the user for input
user_input=take_command()

# Add the user input to the document
selection = word.Selection
selection.TypeText(user_input)

# Save and close the document
dn=input("Enter File Name :")
doc.SaveAs(dn+".docx")
doc.Close()

# Quit Microsoft Word
word.Quit()

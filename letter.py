import openai
from docx import Document

# Configure OpenAI API
openai.api_key = 'YOUR OPEN AI KEY'

def generate_letter(prompt):
    # Make a request to the OpenAI API
    response = openai.Completion.create(
        engine='text-davinci-003',
        prompt=prompt,
        max_tokens=500,
        temperature=0.7,
        n = 1,
        stop=None
    )
    return response.choices[0].text.strip()

def save_as_word_file(content, filename):
    document = Document()
    document.add_paragraph(content)
    document.save(filename)

# Example usage
user_prompt = input("Enter the prompt for the letter: ")
letter_content = generate_letter(user_prompt)
output_filename = 'letter.docx'

save_as_word_file(letter_content, output_filename)
print(f"Letter generated and saved as '{output_filename}'.")

from intelligent_model.prompts.load_prompt import load_prompt
from intelligent_model.services.gemini_service import ask_gemini


def gemini_feedback(docx_content):


    if docx_content:

        # Generate prompt with doc content
        prompt = load_prompt("docx_feedback", documento_text = docx_content)

        try:
            # Send prompt gemini
            response = ask_gemini(prompt)
            return response

        except Exception as ex:

            print("Error in connection with gemini")





# Overview

This script integrates OpenAI's GPT model to process and refine DOCX documents. It reads text from DOCX files, refines the text using GPT, and then saves the updated content to a new DOCX file with appropriate formatting. Here is a detailed description of the main functions and their roles, including an explanation of the iterative refinement process and reasons to continue or stop.

## Main Components

1. **Initialization**  
   - Sets up the OpenAI client.
   - Initializes variables to store conversation history and the finish reason.

2. **Reading DOCX Files**
   - Reads the content of DOCX files and concatenates the text from all paragraphs.

3. **Adding Markdown Paragraphs**
   - Adds paragraphs to a DOCX document with markdown-like formatting (e.g., headings, lists).

4. **Saving to DOCX**
   - Clears the content of a template DOCX file.
   - Adds formatted text (with markdown-like formatting) to the template and saves it as a new DOCX file.

5. **Refining Text with GPT**
   - Uses OpenAI's GPT to ensure the coherence and integration of the updated contract and annex.
   - Iteratively refines the text based on user instructions and GPT responses.

### Steps

1. **Read DOCX Files**  
   - The script reads the original contract and annex DOCX files, extracting and combining their text.

2. **Refine Text with GPT**  
   - The text is processed by GPT to ensure coherence and integration. The model replaces existing sentences in the contract with the corresponding new sentences from the annex, maintaining structure and readability.

3. **Save Refined Text**  
   - The refined text is saved into a new DOCX file, created from a provided template. The text includes appropriate formatting, such as headings and lists, for better readability.

### Refining Text with GPT

The core functionality of refining text using OpenAI's GPT model involves an iterative process where the model integrates new sentences from an annex into an original contract. This ensures that the final document is coherent, well-structured, and readable. Here's a detailed explanation of this process, including the reasons to continue or stop iterations:

#### Iterative Refinement Process

1. **Initial Setup**
   - The process begins by preparing a prompt that contains the original contract, the annex with changes, and clear instructions for the GPT model. This prompt is appended to the conversation history.

2. **Initial Request**
   - The script sends the initial prompt to the GPT model. This prompt includes instructions to identify and replace sentences in the original contract with those from the annex, ensuring the document remains coherent and properly formatted.

3. **Processing Response**
   - The GPT model processes the request and returns a response containing the updated text. This response is appended to the conversation history for context in subsequent iterations.

4. **Handling Finish Reasons**
   - The script checks the `finish_reason` attribute of the GPT model's response to determine whether to continue or stop the refinement process
     - **Continue (finish_reason == "length")** If the response is truncated due to length limitations, the script sends a follow-up prompt asking the model to continue from where it left off. The annex text is omitted in this follow-up to avoid redundancy.
     - **Stop (finish_reason == "stop")** If the response is complete and coherent, the process stops, and the final integrated text is compiled.

5. **Aggregation**
   - Throughout the iterations, the script aggregates the refined text from the model's responses, ensuring that each part seamlessly follows the previous one.

6. **Final Output**
   - Once the model has processed the entire document and no further continuation is required, the aggregated text is ready for saving to a DOCX file.

#### Reasons to Continue or Stop

- **Continue**
  - The script will instruct the GPT model to continue generating text if the response is incomplete due to the maximum token limit. This is indicated by the `finish_reason` being "length". The follow-up prompt will ask the model to continue without repeating the annex text.

- **Stop**
  - The process stops when the GPT model's response is complete, indicated by the `finish_reason` being "stop". At this point, the entire text has been refined and integrated, and no further iterations are needed.

By following this iterative refinement process, the script ensures that the final document is comprehensive and maintains the intended structure and coherence. This approach leverages the strengths of GPT in natural language processing to produce high-quality, readable documents.

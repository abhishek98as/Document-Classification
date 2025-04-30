import google.generativeai as genai
from PIL import Image
import io
import json
import re
import base64
import os

# Configure the Generative AI SDK with your API key
genai.configure(api_key='tesr')

# Initialize the Gemini 2.0 Flash model
model = genai.GenerativeModel('gemini-2.0-flash')

def encode_image_to_base64(image_path):
    """
    Encode an image to base64 string
    
    Args:
        image_path (str): Path to the image file
        
    Returns:
        str: Base64 encoded string of the image
    """
    with Image.open(image_path) as img:
        # Convert to RGB if needed
        if img.mode != 'RGB':
            img = img.convert('RGB')
            
        # Resize if too large
        max_size = (1600, 1600)
        if img.width > max_size[0] or img.height > max_size[1]:
            img.thumbnail(max_size, Image.LANCZOS)
            
        buffered = io.BytesIO()
        img.save(buffered, format='JPEG', quality=95)
        return base64.b64encode(buffered.getvalue()).decode('utf-8')

def extract_text_from_image(image_path, document_type):
    """
    Extracts structured text from an image using the Gemini 2.0 Flash model and formats it as JSON.

    Args:
        image_path (str): Path to the image file.
        document_type (str): Type of document being processed (general description).

    Returns:
        dict: Structured data extracted from the image, or None if an error occurred.
    """
    try:
        # Verify the image file exists
        if not os.path.exists(image_path):
            print(f"Error: Image file not found at {image_path}")
            return None
        
        # Encode image to base64
        base64_image = encode_image_to_base64(image_path)
        
        # Create a general prompt that asks for key-value pairs
        prompt_text = (
            f"Analyze this {document_type} image and extract all important information as key-value pairs.\n\n"
            "Instructions:\n"
            "1. Identify all important fields and their values in the document\n"
            "2. Return ONLY a valid JSON object with keys being the field names and values being the field values\n"
            "3. Use clear, descriptive keys for each field\n"
            "4. Format dates consistently (DD/MM/YYYY where possible)\n"
            "5. If a numeric value is found, preserve its original format\n"
            "6. Do not include any explanations or text outside the JSON object\n"
        )

        # Prepare the image part of the prompt
        image_part = {
            "mime_type": "image/jpeg", 
            "data": base64_image
        }
        
        # Generate content using the model with explicit JSON response format
        try:
            response = model.generate_content(
                contents=[prompt_text, image_part],
                generation_config=genai.GenerationConfig(
                    temperature=0.1,
                    response_mime_type="application/json"
                )
            )
            
            # Get the text from the response
            response_text = response.text.strip()
            
            # Clean up the response to extract JSON
            try:
                # Try to find JSON content within the response using a regex pattern
                json_match = re.search(r'({[\s\S]*})', response_text, re.DOTALL)
                if json_match:
                    response_text = json_match.group(1)
                
                # Parse JSON
                structured_data = json.loads(response_text)
                return structured_data
            except json.JSONDecodeError as json_err:
                print(f"JSON parsing error: {json_err}")
                print(f"Response text: {response_text[:200]}...")
                return try_alternative_extraction(image_path, document_type, base64_image)
                
        except Exception as api_error:
            print(f"API Error: {api_error}")
            return try_alternative_extraction(image_path, document_type, base64_image)

    except Exception as e:
        print(f"An error occurred during image processing: {e}")
        return None

def try_alternative_extraction(image_path, document_type, base64_image=None):
    """
    Alternative approach if the main method fails
    
    Args:
        image_path (str): Path to the image file
        document_type (str): Type of document
        base64_image (str, optional): Base64 encoded image if already available
        
    Returns:
        dict: Extracted data in dictionary format
    """
    try:
        print("Attempting alternative extraction method...")
        
        # Use base64_image if provided, otherwise encode again
        if not base64_image:
            base64_image = encode_image_to_base64(image_path)
        
        # Use a simpler prompt for general key-value extraction
        simple_prompt = (
            "Extract all text from this document image and format each piece of information as key-value pairs. "
            "For each piece of information, identify what type of data it is (name, ID number, date, etc.) "
            "and return it in a format like 'field_name: value'."
        )
        
        # Try with different model parameters
        response = model.generate_content(
            contents=[
                simple_prompt,
                {"mime_type": "image/jpeg", "data": base64_image}
            ],
            generation_config=genai.GenerationConfig(
                temperature=0,
                response_mime_type="text/plain"
            )
        )
        
        extracted_text = response.text
        
        # Try to extract key-value pairs from text
        result = {"raw_text": extracted_text}
        
        # Basic extraction of key-value pairs from text
        pairs = {}
        lines = extracted_text.split('\n')
        for line in lines:
            line = line.strip()
            if ':' in line:
                parts = line.split(':', 1)
                key = parts[0].strip().lower().replace(' ', '_')
                value = parts[1].strip()
                if key and value:
                    pairs[key] = value
        
        if pairs:
            result["extracted_fields"] = pairs
            
        return result
        
    except Exception as e:
        print(f"Alternative extraction failed: {e}")
        return {
            "document_type": document_type,
            "extraction_method": "failed",
            "error": str(e)
        }

def process_document(image_path, document_type):
    """Process a document image and extract information"""
    print(f"Processing {document_type} from {image_path}")
    result = extract_text_from_image(image_path, document_type)
    
    if result:
        print("Extraction result:")
        print(json.dumps(result, indent=2, ensure_ascii=False))
    else:
        print("Failed to extract data from the document.")
    
    return result

# Example usage
if __name__ == "__main__":
    image_path = r"C:\Users\asingh50\Documents\OCR\img1.jpg"
    document_type = "pan card"  # Change this based on your document
    
    # Process the document
    structured_data = process_document(image_path, document_type)

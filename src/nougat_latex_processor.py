# Save this as nougat_latex_processor.py

import os
import torch
from PIL import Image
from transformers import VisionEncoderDecoderModel
from transformers.models.nougat import NougatTokenizerFast
from nougat_latex.util import process_raw_latex_code
from nougat_latex import NougatLaTexProcessor


def process_images_to_latex(
    img_dir, pretrained_model_name_or_path="Norm/nougat-latex-base", device_type="gpu"
):
    # Set device
    device = torch.device("cuda:0" if device_type == "gpu" else "cpu")

    # Initialize model
    model = VisionEncoderDecoderModel.from_pretrained(pretrained_model_name_or_path).to(
        device
    )

    # Initialize processor
    tokenizer = NougatTokenizerFast.from_pretrained(pretrained_model_name_or_path)
    latex_processor = NougatLaTexProcessor.from_pretrained(
        pretrained_model_name_or_path
    )

    results = []

    # Process all images in the specified directory
    for filename in os.listdir(img_dir):
        if filename.endswith((".png", ".jpg", ".jpeg", ".bmp", ".gif")):
            img_path = os.path.join(img_dir, filename)
            image = Image.open(img_path)
            if not image.mode == "RGB":
                image = image.convert("RGB")

            # Process image
            pixel_values = latex_processor(image, return_tensors="pt").pixel_values
            task_prompt = tokenizer.bos_token
            decoder_input_ids = tokenizer(
                task_prompt, add_special_tokens=False, return_tensors="pt"
            ).input_ids

            # Generate LaTeX code
            with torch.no_grad():
                outputs = model.generate(
                    pixel_values.to(device),
                    decoder_input_ids=decoder_input_ids.to(device),
                    max_length=model.config.decoder.max_length,
                    early_stopping=True,
                    pad_token_id=tokenizer.pad_token_id,
                    eos_token_id=tokenizer.eos_token_id,
                    use_cache=True,
                    num_beams=1,
                    bad_words_ids=[[tokenizer.unk_token_id]],
                    return_dict_in_generate=True,
                )

            # Decode and process the output
            sequence = tokenizer.batch_decode(outputs.sequences)[0]
            sequence = (
                sequence.replace(tokenizer.eos_token, "")
                .replace(tokenizer.pad_token, "")
                .replace(tokenizer.bos_token, "")
            )
            sequence = process_raw_latex_code(sequence)
            results.append(sequence)

    return results


# This allows the function to be called directly or imported elsewhere
if __name__ == "__main__":
    # Example usage:
    results = process_images_to_latex(
        img_dir="/home/yanghang/projects/nougat-latex-ocr/file"
    )
    for result in results:
        print(result)

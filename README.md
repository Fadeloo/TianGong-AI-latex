# Env Preparing
Setup `venv`:

```bash
python3.11 -m venv .venv
source .venv/bin/activate
```

Install requirements:

```bash
python.exe -m pip install --upgrade pip

pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
pip install -r requirements.txt --upgrade
```
Test Cuda (optional):

```bash
nvidia-smi
```
# Nougat-LaTeX-OCR

<img src="./asset/img2latex.jpeg" width="600">

Nougat-LaTeX-based is fine-tuned from [facebook/nougat-base](https://huggingface.co/facebook/nougat-base) with [im2latex-100k](https://zenodo.org/record/56198#.V2px0jXT6eA) to boost its proficiency in generating LaTeX code from images. 
Since the initial encoder input image size of nougat was unsuitable for equation image segments, leading to potential rescaling artifacts that degrades the generation quality of LaTeX code. To address this, Nougat-LaTeX-based adjusts the input resolution and uses an adaptive padding approach to ensure that equation image segments in the wild are resized to closely match the resolution of the training data.
Download the model [here](https://huggingface.co/Norm/nougat-latex-base) 👈🏻.

## Uses
### fine-tune on your customized dataset
1. Prepare your dataset in [this](https://drive.google.com/drive/folders/13CA4vAmOmD_I_dSbvLp-Lf0s6KiaNfuO) format
2. Change ``config/base.yaml``
3. Run the training script
```python
python tools/train_experiment.py --config_file config/base.yaml --phase 'train'
```

### use it directly
#### use a pipeline as a high-level helper
```bash
from transformers import pipeline
pipe = pipeline("image-to-text", model="Norm/nougat-latex-base")
```
#### load model directly
```bash
from transformers import AutoTokenizer, AutoModel
tokenizer = AutoTokenizer.from_pretrained("Norm/nougat-latex-base")
model = AutoModel.from_pretrained("Norm/nougat-latex-base")
```

Check the model application information in https://huggingface.co/Norm/nougat-latex-base 
You can find an example in examples_for_latex folder
```python
python examples/run_latex_ocr.py --img_path "examples/test_data/eq1.png"
```


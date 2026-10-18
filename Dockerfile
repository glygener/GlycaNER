FROM python:3.10-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    NLTK_DATA=/usr/local/share/nltk_data \
    HF_HOME=/tmp/hf \
    HOME=/tmp

WORKDIR /app

# CPU-only PyTorch (2.1.0 segfaults on linux/arm64; 2.3.1 gives identical predictions)
RUN pip install torch==2.3.1 \
        --index-url https://download.pytorch.org/whl/cpu \
        --extra-index-url https://pypi.org/simple

COPY requirements.txt /app/requirements.txt
RUN pip install -r /app/requirements.txt

# Constituency parser used in post-processing (1_base_np.py, 3_merge_np.py)
RUN python -c "import benepar; benepar.download('benepar_en3', download_dir='${NLTK_DATA}')"

# Fine-tuned BioBERT model (stored with Git LFS; run `git lfs pull` before building)
COPY model/ /app/model/
RUN python -c "import os,sys; s=os.path.getsize('/app/model/pytorch_model.bin'); sys.exit(0 if s>1e8 else 'model/pytorch_model.bin is a Git LFS pointer: run git lfs pull')"

COPY Scripts/ /app/Scripts/
COPY ["Dictionary Components/", "/app/Scripts/Post_Processing/resources/Dict_Normalization/"]
COPY examples/pmids_example.txt /app/examples/pmids_example.txt
RUN chmod +x /app/Scripts/entrypoint.sh /app/Scripts/Post_Processing/run_pipeline.sh

ENV GLYCANER_MODEL_DIR=/app/model \
    GLYCANER_RESOURCES=/app/Scripts/Post_Processing/resources \
    GLYCANER_PMID_FILE=/data/input/pmids.txt \
    GLYCANER_OUTPUT=/data/output

VOLUME ["/data/output"]
ENTRYPOINT ["/app/Scripts/entrypoint.sh"]
CMD ["all"]

# GlycaNER

GlycaNER is an automated tool that detects and normalizes glycan structure terms
(GSTs) in biomedical literature. A BioBERT model, fine-tuned on 5,200 sentences
from 601 PubMed abstracts (4,725 expert-reviewed terms), finds candidate terms.
Post-processing then filters non-glyco terms, resolves conjunctions and
parentheses, and assigns each term a type: GST, glycoconjugate (GCT), sugar,
sequence, composition or generic. Finally, each term is normalized against the
[Glycan Structure Dictionary (GSD)](https://wiki.glygen.org/Glycan_structure_dictionary)
using the glycan structural components in `Dictionary Components/`.

## Run with Docker

```bash
docker pull shovan5795/glycaner:latest
```

```bash
docker run --rm -v "$PWD/pmids.txt:/data/input/pmids.txt:ro" -v "$PWD/glycaner_output:/data/output" shovan5795/glycaner:latest
```

**Input:** `pmids.txt`, a text file with one PubMed ID per line
(see [`examples/pmids_example.txt`](examples/pmids_example.txt)). Abstracts are
downloaded from PubTator, so the run needs internet access.

**Output:** `glycaner_output/final/` contains the detected glyco-terms for each
sentence, with their type and normalized GSD match (`9c_normalized_form.xlsx`),
plus an HTML view of the abstracts with the detected terms highlighted. See
[`examples/output_45_abstracts`](examples/output_45_abstracts) for the output of
the 45 example PMIDs.

## Repository contents

| Path | Content |
|---|---|
| `Scripts/BERT_Detection/` | PubTator download, sentence splitting and BioBERT term detection |
| `Scripts/Post_Processing/` | term boundary correction, classification and GSD normalization (`run_pipeline.sh` gives the order) |
| `model/` | fine-tuned BioBERT model (Git LFS) |
| `Dictionary Components/` | glycan structural components used for extension and normalization |
| `Data/` | annotated training dataset and evaluation set |
| `Outputs/` | GlycaNER output for 10,000 PubMed abstracts |
| `Dockerfile`, `requirements.txt` | image definition and Python dependencies |

## Citation

If you use GlycaNER, please cite:

> Bhowmik S, Ranzinger R, Au Yeung CCH, Vora J, Bhuiyan U, Tiemeyer M, Mazumder R,
> Vijay-Shanker K. GlycaNER: An Automated Glycan Structure Term Detection and
> Normalization Tool in Biomedical Literature. Manuscript in preparation.

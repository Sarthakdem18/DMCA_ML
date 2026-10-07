# DMCA — Deep Margin Cosine Autoencoder (implementation)

A from-scratch implementation of:

> Wang, Xu, Wang & Xing, "Deep Margin Cosine Autoencoder-Based Medical
> Hyperspectral Image Classification for Tumor Diagnosis," *IEEE
> Transactions on Instrumentation and Measurement*, Vol. 72, 2023.

## What this is

The paper's core idea: a plain autoencoder learns features to reconstruct
its input well, with **no knowledge of the classification labels at all**
(it's unsupervised). A softmax classifier is then bolted on top, but it's
stuck working with features that were never optimized for classification
— they were optimized for reconstruction. DMCA fixes this by folding a
**margin cosine softmax classifier loss directly into the deepest
autoencoder layer's training objective**, so the deepest features are
shaped by both goals (reconstruct well *and* be easy to classify) at once.

## How the code maps to the paper

| File | Paper section / equation |
|---|---|
| `dmca/model.py` | The stacked deep autoencoder (DAE), `H_j - P_j W_j^T H_j` in Eq. (6) |
| `dmca/losses.py` | `ClassifierHead` = Eq. (1)-(3): plain softmax, cosine softmax, margin cosine softmax. `sparsity_loss` = the `‖A‖₁` term in Eq. (8) |
| `dmca/train.py` | `pretrain_stage1` / `train_stage2` = the two-stage training strategy, Section II-C |
| `dmca/data.py` | Synthetic stand-in for the paper's private In Vivo HHB / MC clinical datasets |
| `dmca/metrics.py` | OA / AA / Kappa, Section III-B |
| `main.py` | Reproduces the ablation study structure of Section III-E / Table V |

## Deliberate simplifications (read this before citing results)

1. **No real medical data.** The paper's datasets (In Vivo HHB brain scans,
   MC choledoch scans) are private clinical data, not public. `dmca/data.py`
   generates a synthetic hyperspectral cube with the same structure (4
   classes: background, normal tissue, tumor tissue, blood vessel;
   realistic class imbalance; tumor/normal tissue sharing a common base
   spectrum with only a subtle perturbation between them, like real tissue
   spectra do) so the pipeline and the ablation comparison are meaningful,
   but **the numbers below are not the paper's reported numbers and
   shouldn't be cited as such.**

2. **Optimizer.** The paper solves its combined loss (Eq. 8) with a
   hand-derived alternating-least-squares algorithm (Algorithm 1):
   closed-form update for the decoder `P_J`, per-class gradient descent for
   the encoder/classifier weights `W_J`, soft-thresholding for the sparsity
   auxiliary variable. This code instead minimizes the *identical loss
   function* with a single Adam optimizer via autograd. Same objective,
   different (simpler, more standard) optimization method.

3. **Scale factor `s`.** The paper defines `s = ‖W_k‖₂‖x‖₂` "after"
   normalizing both vectors to unit length — which is mathematically just
   `1`, so this is really informal notation for a fixed rescaling
   hyperparameter, standard in the margin-softmax literature (CosFace/
   ArcFace) this paper borrows from. We treat `s` as a tunable
   hyperparameter (`--s`, default 20) rather than literally recomputing it.

4. **8 baselines → 1.** The paper compares against 8 methods (SVM,
   SVM-Gabor, MFASR, SAE, CNN, RPNet, FMDM, FUST). This code includes one
   representative classical baseline (SVM on raw spectra) plus the 3-rung
   ablation ladder that *is* the paper's actual scientific contribution —
   reimplementing 7 more unrelated architectures wasn't the point of this
   exercise.

## The ablation ladder (the actual point of this code)

This mirrors the paper's own Section III-E / Table V ablation study:

- **`DMCA_NL`** — plain softmax classifier (no cosine normalization, no
  margin). This is "DMCA with the margin-cosine-softmax loss removed."
- **`DMCA_NC`** — cosine softmax, no margin.
- **`DMCA` (full)** — margin cosine softmax. The paper's actual method.

## Running it

```bash
pip install -r requirements.txt
python main.py
```

Takes about a minute on CPU. Prints an OA/AA/Kappa table for all three
ablation variants plus the SVM baseline, and saves two plots to `outputs/`:

- `loss_curve.png` — Stage-2 training loss over iterations (cf. paper Fig. 3)
- `classification_maps.png` — ground truth vs. SVM vs. full-DMCA pixel maps
  (cf. paper Figs. 5-8)

## What the results actually show

With the default synthetic data, `tumor_tissue` is deliberately the hardest
class (its spectrum is a subtle variant of `normal_tissue`'s, buried in
correlated noise) — this is where the paper's motivation should visibly
matter. In a representative run:

| Method | OA | AA | Kappa | Tumor recall |
|---|---|---|---|---|
| DMCA_NL (plain softmax) | 91.2% | 75.5% | 0.848 | 2.1% |
| DMCA_NC (cosine, no margin) | 94.7% | 88.4% | 0.911 | 57.4% |
| DMCA (full, margin cosine) | 94.5% | 89.1% | 0.909 | 61.5% |

The plain softmax classifier essentially gives up on the minority/hard
tumor class (2% recall) — a textbook symptom of features that weren't
shaped with the classifier in mind. Moving to a cosine-normalized softmax,
and then adding the margin, recovers tumor recall dramatically and gives
the best **AA** (average accuracy, the metric that doesn't get dominated by
the easy majority classes) of the three — the same ordering
(`DMCA > DMCA_NC > DMCA_NL`) the paper's own Table V reports. Numbers will
vary run-to-run with the random seed / hyperparameters (`--help` isn't
wired up yet — edit the constants at the top of `main.py` to experiment:
`lambda1`, `lambda2`, `s`, `m`, `dims`, `train_ratio`).

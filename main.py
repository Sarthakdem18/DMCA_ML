"""
End-to-end demo of the DMCA method (Wang, Xu, Wang & Xing, 2023).

Runs:
  1. Generate a synthetic MedHSI cube (stand-in for the paper's private
     clinical datasets -- see dmca/data.py for why).
  2. Train the 3-rung classifier-head ablation ladder described in the
     paper's own ablation study (Section III-E, Table V):
         DMCA_NL       -> plain softmax            (no cosine, no margin)
         DMCA_NC       -> cosine softmax            (no margin)
         DMCA (full)   -> margin cosine softmax     (paper's method)
  3. Train an SVM-on-raw-pixels baseline, mirroring the paper's weakest
     baseline (Table I/II).
  4. Print an OA/AA/Kappa comparison table (paper-style).
  5. Save classification-map and training-loss plots to outputs/.

Run with:  python main.py
"""
import numpy as np
import torch
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from sklearn.svm import SVC

from dmca.data import generate_medhsi_cube, flatten_cube, stratified_split, CLASS_NAMES, N_CLASSES
from dmca.train import train_dmca, predict
from dmca.metrics import summarize, format_summary

SEED = 0
torch.manual_seed(SEED)
np.random.seed(SEED)

OUT_DIR = "outputs"


def main():
    H, W, T = 100, 100, 60
    print(f"Generating synthetic MedHSI cube: {H}x{W} pixels, {T} spectral bands, "
          f"{N_CLASSES} classes {CLASS_NAMES}")
    cube, label = generate_medhsi_cube(H=H, W=W, T=T, seed=SEED)
    X, y = flatten_cube(cube, label)

    train_idx, test_idx = stratified_split(y, train_ratio=0.08, seed=SEED, min_train_per_class=25)
    print(f"Train pixels: {len(train_idx)}   Test pixels: {len(test_idx)}")
    for c, name in enumerate(CLASS_NAMES):
        print(f"  class {c} ({name:>13s}): {np.sum(y[train_idx]==c):4d} train / {np.sum(y[test_idx]==c):5d} test")

    X_train_t = torch.tensor(X[train_idx], dtype=torch.float32)
    y_train_t = torch.tensor(y[train_idx], dtype=torch.long)
    X_test_t = torch.tensor(X[test_idx], dtype=torch.float32)
    y_test_np = y[test_idx]

    # Network shape: T -> 32 -> 16 -> 8   (J=3 encode layers, matching the
    # paper's "number of deep layers J+1 fixed to 4" once you count the
    # classifier head as the +1).
    dims = [T, 32, 16, 8]
    lambda1, lambda2 = 0.1, 0.01  # the paper's own grid-search optimum (Fig. 4)

    ablation_modes = [
        ("DMCA_NL (plain softmax)", "plain"),
        ("DMCA_NC (cosine, no margin)", "cosine"),
        ("DMCA (full, margin cosine)", "margin_cosine"),
    ]

    results = {}
    trained = {}
    print("\n=== Training ablation ladder ===")
    for label_name, mode in ablation_modes:
        print(f"\n-- {label_name} --")
        encoder, head, history = train_dmca(
            X_train_t, y_train_t, dims, N_CLASSES, mode=mode,
            lambda1=lambda1, lambda2=lambda2, s=20.0, m=0.25,
            epochs_stage1=300, epochs_stage2=800, lr=1e-2, verbose=True,
        )
        y_pred = predict(encoder, head, X_test_t).numpy()
        res = summarize(y_test_np, y_pred, N_CLASSES, CLASS_NAMES)
        results[label_name] = res
        trained[label_name] = (encoder, head, history)

    # --- SVM baseline on raw spectral pixels (paper's weakest baseline) ---
    print("\n-- SVM (raw spectra) baseline --")
    svm = SVC(kernel="rbf", C=10, gamma="scale")
    svm.fit(X[train_idx], y[train_idx])
    y_pred_svm = svm.predict(X[test_idx])
    results["SVM (raw spectra)"] = summarize(y_test_np, y_pred_svm, N_CLASSES, CLASS_NAMES)

    # --- Results table ---
    print("\n=== Results (paper-style OA / AA / Kappa comparison) ===")
    for name, res in results.items():
        print(format_summary(name, res))

    import os
    os.makedirs(OUT_DIR, exist_ok=True)

    # --- Stage-2 loss curve for the full DMCA model (mirrors paper Fig. 3) ---
    _, _, full_history = trained["DMCA (full, margin cosine)"]
    plt.figure(figsize=(6, 4))
    plt.plot(full_history)
    plt.xlabel("Stage-2 training iteration")
    plt.ylabel("Combined DMCA loss")
    plt.title("Evolution of the DMCA cost function (cf. paper Fig. 3)")
    plt.tight_layout()
    plt.savefig(f"{OUT_DIR}/loss_curve.png", dpi=150)
    plt.close()

    # --- Classification maps (mirrors paper Figs. 5-8) ---
    X_all_t = torch.tensor(X, dtype=torch.float32)
    encoder_full, head_full, _ = trained["DMCA (full, margin cosine)"]
    y_pred_full_all = predict(encoder_full, head_full, X_all_t).numpy().reshape(H, W)
    y_pred_svm_all = svm.predict(X).reshape(H, W)

    cmap = matplotlib.colors.ListedColormap(["#222222", "#4C72B0", "#C44E52", "#55A868"])
    fig, axes = plt.subplots(1, 3, figsize=(12, 4))
    for ax, data, title in zip(
        axes,
        [label, y_pred_svm_all, y_pred_full_all],
        ["Ground truth", "SVM prediction", "DMCA (full) prediction"],
    ):
        ax.imshow(data, cmap=cmap, vmin=0, vmax=N_CLASSES - 1)
        ax.set_title(title)
        ax.axis("off")
    plt.tight_layout()
    plt.savefig(f"{OUT_DIR}/classification_maps.png", dpi=150)
    plt.close()

    print(f"\nSaved plots to {OUT_DIR}/loss_curve.png and {OUT_DIR}/classification_maps.png")


if __name__ == "__main__":
    main()

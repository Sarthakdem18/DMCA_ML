"""
Classifier heads corresponding to the paper's Eq. (1)-(3), plus the L1
sparsity term used in Eq. (6)/(8).

We implement a 3-rung ladder of classifier heads. This exactly matches the
paper's own ablation study (Section III-E / Table V):

  mode="plain"        -> ordinary softmax classifier, Eq. (1).
                          = "DMCA_NL": DMCA with the margin-cosine-softmax
                          loss removed entirely.
  mode="cosine"        -> soft-max on L2-normalized features/weights
                          (Eq. 2) but with NO margin.
                          = "DMCA_NC": DMCA without the cosine margin.
  mode="margin_cosine" -> the full paper loss, Eq. (3): cosine softmax
                          with an angular margin m subtracted from the
                          true class before the softmax.
                          = full DMCA.

Note on the scale factor s: the paper defines s = ||W_k||_2 * ||x||_2
"before" normalization is applied, which is a slightly informal way of
saying what CosFace/ArcFace-style margin-softmax losses do in practice --
after L2-normalizing both vectors, cosine similarity is squeezed into
[-1, 1], which is too flat a range for cross-entropy gradients to separate
classes well. So s is used as a fixed rescaling hyperparameter (we default
to 30, the standard value used in the margin-softmax literature this
paper borrows from) rather than something literally recomputed from
already-unit-norm vectors. This is a deliberate, documented simplification.
"""
import torch
import torch.nn as nn
import torch.nn.functional as F


class ClassifierHead(nn.Module):
    def __init__(self, feat_dim, n_classes, mode="margin_cosine", s=30.0, m=0.35):
        super().__init__()
        assert mode in ("plain", "cosine", "margin_cosine")
        self.mode = mode
        self.s = s
        self.m = m if mode == "margin_cosine" else 0.0

        if mode == "plain":
            self.linear = nn.Linear(feat_dim, n_classes)
        else:
            # W_J in the paper: one "class prototype" direction per class.
            self.W = nn.Parameter(torch.randn(feat_dim, n_classes) * 0.01)

    def forward(self, x, labels=None):
        """
        x: (batch, feat_dim) deepest DAE features (H_J in the paper).
        labels: (batch,) int64 ground-truth class ids, or None at
                inference time (margin is only ever applied to the *true*
                class, so it needs labels; at test time we don't know the
                true class, so we just return plain scaled cosine logits).
        Returns: (batch, n_classes) logits suitable for F.cross_entropy.
        """
        if self.mode == "plain":
            return self.linear(x)

        x_n = F.normalize(x, dim=1)
        w_n = F.normalize(self.W, dim=0)
        cosine = x_n @ w_n  # cos(W_k, x) for every class k, Eq. (2)

        if labels is None or self.m == 0.0:
            return self.s * cosine

        onehot = F.one_hot(labels, num_classes=cosine.size(1)).float()
        return self.s * (cosine - self.m * onehot)  # Eq. (3)

    @torch.no_grad()
    def predict(self, x):
        logits = self.forward(x, labels=None)
        return logits.argmax(dim=1)


def sparsity_loss(features):
    """L1 penalty on the deepest encoded features, i.e. the ||A||_1 term
    in Eq. (8), which encourages a compact/sparse feature representation."""
    return features.abs().mean()

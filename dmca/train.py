"""
Two-stage training strategy (paper Section II-C).

Stage 1 (unsupervised, greedy layer-wise pretraining):
    Train encoder layers 1..J-1 purely on reconstruction loss, one layer
    at a time, each layer learning to reconstruct the *previous* layer's
    output. No labels are used here at all -- this is the classic "plain
    autoencoder" behavior the paper argues is classifier-agnostic (see the
    "why can't an autoencoder classify" discussion). This initializes a
    reasonable, generic feature extractor before we touch labels.

Stage 2 (supervised, the actual DMCA contribution):
    Freeze layers 1..J-1. Take their output Z = H_{J-1} as fixed input.
    Train ONLY the deepest layer J (its encode W_J and decode P_J) plus
    the classifier head, jointly, on the full combined loss from Eq. (6):

        total_loss = reconstruction_loss(Z, decode(encode(Z)))
                   + lambda1 * sparsity_loss(encode(Z))
                   + lambda2 * classifier_loss(encode(Z), labels)

    This is exactly where "the importance of the classifier is added as a
    constraint into the DAE network" (the paper's core idea) actually
    happens: layer J's features are shaped by BOTH the reconstruction
    objective AND the classification objective simultaneously.

Implementation note: the paper solves Eq. (8) with a hand-derived
alternating-least-squares optimizer (Algorithm 1: closed-form update for
P_J, per-class gradient descent for W_J, soft-thresholding for the
sparsity auxiliary variable A). We instead minimize the identical loss
function with a single Adam optimizer via autograd. Both are valid ways
to minimize the same objective; autograd is simpler to read, verify, and
is how this would actually be built in practice today.
"""
import torch
import torch.nn.functional as F

from .model import AELayer, StackedEncoder
from .losses import ClassifierHead, sparsity_loss


def pretrain_stage1(X_train, dims, epochs=300, lr=1e-2, verbose=False):
    """Greedy layer-wise autoencoder pretraining on reconstruction loss
    only. Returns a StackedEncoder with layers 1..J-1 already trained,
    plus the deepest layer left randomly initialized (Stage 2 trains it)."""
    encoder = StackedEncoder(dims)
    current_input = X_train

    n_outer_layers = len(encoder.layers) - 1  # everything except layer J
    for li in range(n_outer_layers):
        layer = encoder.layers[li]
        opt = torch.optim.Adam(layer.parameters(), lr=lr)
        for epoch in range(epochs):
            opt.zero_grad()
            _, recon = layer(current_input)
            loss = F.mse_loss(recon, current_input)
            loss.backward()
            opt.step()
        if verbose:
            print(f"  [stage1] layer {li+1}/{n_outer_layers} final recon MSE: {loss.item():.5f}")
        with torch.no_grad():
            current_input, _ = layer(current_input)

    return encoder


def train_stage2(encoder, X_train, y_train, n_classes, mode="margin_cosine",
                  lambda1=0.1, lambda2=0.01, s=30.0, m=0.35,
                  epochs=800, lr=1e-2, verbose=False):
    """Trains the deepest layer J + classifier head jointly on the
    combined DMCA loss. `mode` selects which classifier ladder rung is
    used (plain / cosine / margin_cosine) -- this is what the ablation
    study in main.py sweeps over."""
    J = len(encoder.layers)
    Z = encoder.forward_frozen_prefix(X_train, up_to_layer=J - 1)  # H_{J-1}

    layer_J = encoder.layers[J - 1]
    feat_dim = layer_J.encode.out_features
    head = ClassifierHead(feat_dim, n_classes, mode=mode, s=s, m=m)

    params = list(layer_J.parameters()) + list(head.parameters())
    opt = torch.optim.Adam(params, lr=lr)

    history = []
    for epoch in range(epochs):
        opt.zero_grad()
        h, recon = layer_J(Z)
        recon_loss = F.mse_loss(recon, Z)
        sparse_loss = sparsity_loss(h)
        logits = head(h, labels=y_train)
        cls_loss = F.cross_entropy(logits, y_train)

        loss = recon_loss + lambda1 * sparse_loss + lambda2 * cls_loss
        loss.backward()
        opt.step()
        history.append(loss.item())

        if verbose and (epoch % 100 == 0 or epoch == epochs - 1):
            print(f"  [stage2:{mode}] epoch {epoch:4d}  total={loss.item():.4f} "
                  f"recon={recon_loss.item():.4f} sparse={sparse_loss.item():.4f} "
                  f"cls={cls_loss.item():.4f}")

    return head, history


def predict(encoder, head, X):
    """Full forward pass: raw spectra -> layers 1..J -> classifier head."""
    with torch.no_grad():
        z = X
        for layer in encoder.layers:
            z, _ = layer(z)
        return head.predict(z)


def train_dmca(X_train, y_train, dims, n_classes, mode="margin_cosine",
                lambda1=0.1, lambda2=0.01, s=30.0, m=0.35,
                epochs_stage1=300, epochs_stage2=800, lr=1e-2, verbose=False):
    """Convenience wrapper: runs Stage 1 then Stage 2 and returns the
    trained (encoder, classifier_head, stage2_loss_history)."""
    encoder = pretrain_stage1(X_train, dims, epochs=epochs_stage1, lr=lr, verbose=verbose)
    head, history = train_stage2(
        encoder, X_train, y_train, n_classes, mode=mode,
        lambda1=lambda1, lambda2=lambda2, s=s, m=m,
        epochs=epochs_stage2, lr=lr, verbose=verbose,
    )
    return encoder, head, history

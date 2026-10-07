"""
The Deep Autoencoder (DAE) backbone, i.e. the H_j -> P_j W_j^T H_j pieces
of Eq. (6). Each AELayer is one autoencoder "layer": it encodes its input
down to a smaller representation (W_j) and decodes it back (P_j), and is
scored by how well the reconstruction matches the input.

We use J stacked layers. Layers 1..J-1 are the ordinary/"outer" encoder
layers. Layer J is the special deepest layer that, in Stage 2 training
(see train.py), also has to satisfy the margin-cosine-softmax classifier
constraint -- that's the actual DMCA contribution.
"""
import torch
import torch.nn as nn


class AELayer(nn.Module):
    def __init__(self, in_dim, out_dim):
        super().__init__()
        self.encode = nn.Linear(in_dim, out_dim)  # W_j
        self.decode = nn.Linear(out_dim, in_dim)   # P_j
        self.act = nn.Tanh()

    def forward(self, x):
        h = self.act(self.encode(x))
        recon = self.decode(h)
        return h, recon


class StackedEncoder(nn.Module):
    """Holds every AELayer 1..J. dims = [T, d1, d2, ..., feat_dim] describes
    T (input spectral bands) down to the final feature dimension used by
    the classifier."""

    def __init__(self, dims):
        super().__init__()
        self.layers = nn.ModuleList(
            [AELayer(dims[i], dims[i + 1]) for i in range(len(dims) - 1)]
        )

    def forward_frozen_prefix(self, x, up_to_layer):
        """Runs x through layers[0:up_to_layer] with no gradient tracking
        -- this produces H_{J-1}, the input that the special Jth layer
        receives, matching the paper's Stage 1 -> Stage 2 hand-off."""
        with torch.no_grad():
            for layer in self.layers[:up_to_layer]:
                x, _ = layer(x)
        return x

    def forward_full(self, x):
        """Runs x through every layer with gradients (used only for
        end-to-end evaluation/prediction after training is complete)."""
        for layer in self.layers:
            x, _ = layer(x)
        return x

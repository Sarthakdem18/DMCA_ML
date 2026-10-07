"""
Synthetic medical-hyperspectral-image (MedHSI) generator.

The paper trains on two real, non-public clinical datasets (In Vivo HHB
brain-cancer scans, and the MC choledoch/bile-duct dataset). We don't have
access to those, so this module fabricates a stand-in dataset with the same
*structure* the paper describes:

  - a 3-D cube  H x W x T  (spatial height x width, T spectral bands)
  - a handful of classes with distinct, noisy spectral "fingerprints"
  - realistic class imbalance (background dominates; a thin blood-vessel
    class is a tiny sliver of pixels), matching the four classes used in
    the paper's In Vivo HHB dataset: normal tissue, tumor tissue, blood
    vessel, background.

This lets every part of the DMCA pipeline (training, evaluation, ablation,
classification maps) be exercised end-to-end and gives a fair, reproducible
apples-to-apples comparison between the plain/cosine/margin-cosine
classifier heads -- which is the actual scientific question the paper asks.
It does NOT reproduce the paper's reported numbers, since it isn't the same
data.
"""
import numpy as np

CLASS_NAMES = ["background", "normal_tissue", "tumor_tissue", "blood_vessel"]
N_CLASSES = len(CLASS_NAMES)


def _spectral_signature(rng, T, n_bumps=3, baseline=0.2):
    """A smooth, class-specific reflectance curve = a few Gaussian bumps
    at random band positions, plus a flat baseline. Different classes get
    different bump positions/heights, mimicking how different tissue types
    reflect light differently at different wavelengths."""
    x = np.linspace(0, 1, T)
    sig = np.full(T, baseline, dtype=np.float64)
    for _ in range(n_bumps):
        center = rng.uniform(0.05, 0.95)
        width = rng.uniform(0.03, 0.12)
        height = rng.uniform(0.3, 1.0)
        sig += height * np.exp(-((x - center) ** 2) / (2 * width ** 2))
    return sig


def _perturb_signature(rng, base, T, height_range=(0.12, 0.20), width_range=(0.02, 0.05)):
    """Adds ONE small, subtle bump on top of an existing signature. Used to
    make 'normal tissue' and 'tumor tissue' close variants of a shared
    tissue spectrum rather than wildly different curves -- real tumor vs.
    healthy tissue spectra are subtly different, not night-and-day
    different, which is what makes the classification task genuinely
    hard (and is why this demo is interesting)."""
    x = np.linspace(0, 1, T)
    center = rng.uniform(0.05, 0.95)
    width = rng.uniform(*width_range)
    height = rng.uniform(*height_range)
    bump = height * np.exp(-((x - center) ** 2) / (2 * width ** 2))
    return base + bump


def _band_smoothing_matrix(T, sigma):
    """A (T, T) matrix that, applied to i.i.d. noise, produces noise that
    is correlated across neighboring spectral bands (real hyperspectral
    sensor noise isn't independent per-band -- neighboring wavelengths are
    correlated). This is what makes the classification problem non-trivial:
    smooth, broad noise fluctuations can masquerade as the same kind of
    subtle bump used to distinguish normal vs. tumor tissue above."""
    idx = np.arange(T)
    diff = idx[:, None] - idx[None, :]
    K = np.exp(-(diff ** 2) / (2 * sigma ** 2))
    K /= K.sum(axis=1, keepdims=True)
    return K


def _ellipse_mask(H, W, cy, cx, ry, rx):
    yy, xx = np.mgrid[0:H, 0:W]
    return (((yy - cy) / ry) ** 2 + ((xx - cx) / rx) ** 2) <= 1.0


def _vessel_mask(H, W, rng, thickness=1.5):
    """A thin, wandering blood-vessel-like curve across the image."""
    mask = np.zeros((H, W), dtype=bool)
    y = rng.uniform(H * 0.2, H * 0.8)
    for x in range(W):
        y += rng.normal(0, H * 0.03)
        y = np.clip(y, 2, H - 3)
        y0, y1 = int(y - thickness), int(y + thickness)
        mask[max(0, y0):min(H, y1 + 1), x] = True
    return mask


def generate_medhsi_cube(H=64, W=64, T=60, noise_std=0.10, band_corr_sigma=2.0, seed=0):
    """
    Returns:
        cube  : (H, W, T) float32 array of synthetic hyperspectral pixels
        label : (H, W) int array with values in [0, N_CLASSES)
    """
    rng = np.random.default_rng(seed)

    # Background and blood vessel get their own distinct spectra (easy to
    # tell apart). Normal and tumor tissue SHARE a base "tissue" spectrum
    # and only differ by one small, subtle bump each -- this is the hard,
    # medically-realistic part of the task, and is exactly where a better
    # feature extractor (DMCA) should have an edge over a weaker one.
    background_sig = _spectral_signature(rng, T, n_bumps=2, baseline=0.15)
    vessel_sig = _spectral_signature(rng, T, n_bumps=2, baseline=0.5)
    base_tissue_sig = _spectral_signature(rng, T, n_bumps=2, baseline=0.25)
    normal_sig = _perturb_signature(rng, base_tissue_sig, T)
    tumor_sig = _perturb_signature(rng, base_tissue_sig, T)

    signatures = np.stack([background_sig, normal_sig, tumor_sig, vessel_sig], axis=0)

    label = np.zeros((H, W), dtype=np.int64)  # start as all "background"

    # A big "normal tissue" region
    label[_ellipse_mask(H, W, cy=H * 0.55, cx=W * 0.45, ry=H * 0.38, rx=W * 0.42)] = 1
    # A "tumor" blob sitting inside/near the normal tissue region
    label[_ellipse_mask(H, W, cy=H * 0.42, cx=W * 0.55, ry=H * 0.16, rx=W * 0.18)] = 2
    # A thin blood-vessel streak
    label[_vessel_mask(H, W, rng)] = 3

    cube = signatures[label]  # (H, W, T) via fancy indexing

    # Spectrally-correlated noise (see _band_smoothing_matrix docstring).
    K = _band_smoothing_matrix(T, sigma=band_corr_sigma)
    raw_noise = rng.normal(0, 1, size=(H * W, T))
    smooth_noise = raw_noise @ K.T
    smooth_noise *= noise_std / smooth_noise.std()  # normalize to target std
    cube = cube + smooth_noise.reshape(H, W, T)
    cube = cube.astype(np.float32)

    return cube, label


def flatten_cube(cube, label):
    H, W, T = cube.shape
    X = cube.reshape(-1, T)
    y = label.reshape(-1)
    return X, y


def stratified_split(y, train_ratio, seed=0, min_train_per_class=5):
    """Randomly pick a small fraction of pixels per class for training,
    the rest for testing -- mirrors the paper's protocol of training on a
    tiny percentage of labeled pixels (they use 0.1%-0.2%; we use a larger
    ratio by default since our synthetic image is far smaller than a real
    ~1000x1000 clinical scan and needs more samples to learn stably)."""
    rng = np.random.default_rng(seed)
    train_idx, test_idx = [], []
    for c in np.unique(y):
        idx = np.where(y == c)[0]
        rng.shuffle(idx)
        n_train = max(min_train_per_class, int(len(idx) * train_ratio))
        n_train = min(n_train, len(idx) - 1)  # leave at least 1 for testing
        train_idx.append(idx[:n_train])
        test_idx.append(idx[n_train:])
    train_idx = np.concatenate(train_idx)
    test_idx = np.concatenate(test_idx)
    rng.shuffle(train_idx)
    rng.shuffle(test_idx)
    return train_idx, test_idx

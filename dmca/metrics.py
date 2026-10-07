"""
Evaluation metrics used by the paper (Section III-B):
  - OA    (Overall Accuracy):  % of all test pixels labeled correctly
  - AA    (Average Accuracy):  mean of the per-class recall (fair when
                                classes are imbalanced, e.g. "background"
                                pixels vastly outnumber "tumor" pixels)
  - Kappa (Cohen's Kappa):     accuracy corrected for the accuracy you'd
                                get from random guessing
"""
import numpy as np
from sklearn.metrics import accuracy_score, cohen_kappa_score, confusion_matrix


def overall_accuracy(y_true, y_pred):
    return accuracy_score(y_true, y_pred)


def per_class_accuracy(y_true, y_pred, n_classes):
    """Recall of each class: correct-in-class / total-in-class."""
    cm = confusion_matrix(y_true, y_pred, labels=list(range(n_classes)))
    with np.errstate(divide="ignore", invalid="ignore"):
        per_class = np.diag(cm) / cm.sum(axis=1)
    return np.nan_to_num(per_class)


def average_accuracy(y_true, y_pred, n_classes):
    return float(np.mean(per_class_accuracy(y_true, y_pred, n_classes)))


def kappa(y_true, y_pred):
    return cohen_kappa_score(y_true, y_pred)


def summarize(y_true, y_pred, n_classes, class_names=None):
    """Returns a dict matching the style of the paper's result tables."""
    pca = per_class_accuracy(y_true, y_pred, n_classes)
    result = {
        "OA": overall_accuracy(y_true, y_pred),
        "AA": average_accuracy(y_true, y_pred, n_classes),
        "Kappa": kappa(y_true, y_pred),
        "per_class": {
            (class_names[i] if class_names else str(i)): float(pca[i])
            for i in range(n_classes)
        },
    }
    return result


def format_summary(name, result):
    lines = [f"{name:>14s} | OA={result['OA']*100:6.2f}%  AA={result['AA']*100:6.2f}%  Kappa={result['Kappa']:.4f}"]
    for cls, acc in result["per_class"].items():
        lines.append(f"{'':>14s} |   {cls:<14s}: {acc*100:6.2f}%")
    return "\n".join(lines)

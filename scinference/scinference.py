import numpy as np
from scipy.stats import t as t_dist
from conformal import movingblock, iid, confidence_interval
from ttest import sc_cf, did_cf


def scinference(Y1, Y0, T1, T0,
                inference_method="conformal",
                alpha=0.1,
                ci=False,
                theta0=0,
                estimation_method="sc",
                permutation_method="mb",
                ci_grid=None,
                n_perm=5000,
                lsei_type=1,
                K=2):
    """
    Inference methods for synthetic control.

    Parameters:
    -----------
    Y1 : array-like
        Outcome data for treated unit (T x 1 vector)
    Y0 : array-like
        Outcome data for control units (T x J matrix)
    T1 : int
        Number of post-treatment periods
    T0 : int
        Number of pre-treatment periods
    inference_method : str, optional
        Inference method: "conformal" (default) or "ttest"
    alpha : float, optional
        Significance level (default=0.1)
    ci : bool, optional
        Whether pointwise confidence intervals are reported (default=False)
    theta0 : float or array-like, optional
        Null hypothesis for treatment effect trajectory; default=0
    estimation_method : str, optional
        Estimation method: "did", "sc" (default), or "classo"
    permutation_method : str, optional
        Permutation method: "mb" (moving block, default) or "iid"
    ci_grid : array-like, optional
        Grid for the confidence interval
    n_perm : int, optional
        Number of permutations; default=5000
    lsei_type : int, optional
        Option for lsei used for sc; default=1
    K : int, optional
        Number of cross-fits for t-test (K > 1); default=2

    Returns:
    --------
    dict : Dictionary with results
    """
    # Convert inputs to numpy arrays
    Y1 = np.asarray(Y1, dtype=np.float64).flatten()
    Y0 = np.asarray(Y0, dtype=np.float64)

    # Input validation
    if len(Y1) != (T0 + T1):
        raise ValueError("length of Y1 needs to be equal to T0 + T1")
    if Y0.shape[0] != (T0 + T1):
        raise ValueError("number of rows in Y0 needs to be equal to T0 + T1")
    if inference_method not in ["conformal", "ttest"]:
        raise ValueError("The selected inference method is not available")

    # Conformal inference
    if inference_method == "conformal":

        if estimation_method not in ["did", "sc", "classo"]:
            raise ValueError("The selected estimation method is not implemented for conformal inference")
        if permutation_method not in ["iid", "mb"]:
            raise ValueError("The selected class of permutations is not available")

        theta0 = np.asarray(theta0).flatten()
        if len(theta0) != 1 and len(theta0) != T1:
            raise ValueError("length of theta0 should be 1 or T1")

        # P-value
        if permutation_method == "mb":
            p_val = movingblock(
                Y1=Y1, Y0=Y0, T1=T1, T0=T0, 
                theta0=theta0,
                estimation_method=estimation_method, 
                lsei_type=lsei_type
            )
        else:
            p_val = iid(
                Y1=Y1, Y0=Y0, T1=T1, T0=T0, 
                theta0=theta0,
                estimation_method=estimation_method, 
                n_perm=n_perm, 
                lsei_type=lsei_type
            )

        # Confidence intervals
        if ci:
            if ci_grid is None:
                raise ValueError("no grid specified for confidence interval")

            ci_grid = np.asarray(ci_grid)
            obj = confidence_interval(
                Y1=Y1, Y0=Y0, T1=T1, T0=T0,
                estimation_method=estimation_method,
                alpha=alpha, 
                ci_grid=ci_grid, 
                lsei_type=lsei_type
            )
            lb = obj['lb']
            ub = obj['ub']
        else:
            lb = np.nan
            ub = np.nan

        return {'p_val': p_val, 'lb': lb, 'ub': ub}

    # T-test
    elif inference_method == "ttest":

        if estimation_method not in ["did", "sc"]:
            raise ValueError("The selected estimation method is not implemented for the t-test")
        if K <= 1:
            raise ValueError("K must be strictly greater than 1")

        if estimation_method == "did":
            obj = did_cf(Y1, Y0, T1, T0, K)
        else:
            obj = sc_cf(Y1, Y0, T1, T0, K, lsei_type)

        att = obj['tau_hat']
        se = obj['se_hat']

        t_critical = t_dist.ppf(1 - alpha / 2, df=K - 1)
        lb = att - t_critical * se
        ub = att + t_critical * se

        return {'att': att, 'se': se, 'lb': lb, 'ub': ub}
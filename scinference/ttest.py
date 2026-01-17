import numpy as np
from estimators import sc


def sc_cf(Y1, Y0, T1, T0, K, lsei_type=1):
    """
    Synthetic control with cross-fitting t-test.
    
    Equivalent to R's sc.cf function.

    Parameters:
    -----------
    Y1 : numpy.ndarray
        Treated unit outcomes (T01 x 1 vector)
    Y0 : numpy.ndarray
        Control units outcomes (T01 x J matrix)
    T1 : int
        Number of post-treatment periods
    T0 : int
        Number of pre-treatment periods
    K : int
        Number of folds for cross-fitting
    lsei_type : int, optional
        Type of least squares solver for sc method (default=1)

    Returns:
    --------
    dict : Dictionary containing:
        - t_hat: t-statistic (follows t_{K-1} distribution)
        - tau_hat: estimated treatment effect
        - se_hat: standard error
    """
    T01 = T0 + T1
    
    # R: r <- min(floor(T0/K), T1)
    r = min(int(np.floor(T0 / K)), T1)

    Y1_pre = Y1[:T0]
    Y0_pre = Y0[:T0, :]
    Y1_post = Y1[T0:T01]
    Y0_post = Y0[T0:T01, :]

    tau_mat = np.zeros(K)

    for k in range(K):
        # R: Hk <- (T0-(r*K)) + seq((k-1)*r+1, k*r, 1)
        # R is 1-indexed, so we need to adjust
        # In R: for k=1, seq(1, r) gives indices 1 to r
        # In Python: for k=0, we want indices 0 to r-1
        
        # R formula: start = (T0 - r*K) + (k-1)*r + 1 = T0 - r*K + k*r - r + 1 = T0 - r*(K-k+1) + 1
        # R formula: end = (T0 - r*K) + k*r = T0 - r*(K-k)
        
        # Simplified for Python (0-indexed):
        base = T0 - (r * K)
        start_idx = base + k * r
        end_idx = base + (k + 1) * r
        Hk = np.arange(start_idx, end_idx)

        # Training indices (exclude Hk)
        train_idx = np.setdiff1d(np.arange(T0), Hk)

        # Estimate weights on training set
        w_Hk = sc(Y1_pre[train_idx], Y0_pre[train_idx, :], lsei_type)['w_hat']

        # R: tau.mat[k,1] <- mean(Y1.post - Y0.post %*% w.Hk) - mean(Y1.pre[Hk] - Y0.pre[Hk,] %*% w.Hk)
        tau_mat[k] = (np.mean(Y1_post - Y0_post @ w_Hk) - 
                      np.mean(Y1_pre[Hk] - Y0_pre[Hk, :] @ w_Hk))

    # R: tau.hat <- mean(tau.mat)
    tau_hat = np.mean(tau_mat)
    
    # R: se.hat <- sqrt(1 + ((K*r)/T1)) * sd(tau.mat) / sqrt(K)
    # Note: R's sd() uses n-1 denominator (ddof=1)
    se_hat = np.sqrt(1 + ((K * r) / T1)) * np.std(tau_mat, ddof=1) / np.sqrt(K)
    
    # R: t.hat <- tau.hat / se.hat
    t_hat = tau_hat / se_hat

    return {'t_hat': t_hat, 'tau_hat': tau_hat, 'se_hat': se_hat}


def did_cf(Y1, Y0, T1, T0, K):
    """
    Difference-in-differences with cross-fitting t-test.
    
    Equivalent to R's did.cf function.

    Parameters:
    -----------
    Y1 : numpy.ndarray
        Treated unit outcomes (T01 x 1 vector)
    Y0 : numpy.ndarray
        Control units outcomes (T01 x J matrix)
    T1 : int
        Number of post-treatment periods
    T0 : int
        Number of pre-treatment periods
    K : int
        Number of folds for cross-fitting

    Returns:
    --------
    dict : Dictionary containing:
        - t_hat: t-statistic (follows t_{K-1} distribution)
        - tau_hat: estimated treatment effect
        - se_hat: standard error
    """
    T01 = T0 + T1
    r = min(int(np.floor(T0 / K)), T1)

    Y1_pre = Y1[:T0]
    Y0_pre = Y0[:T0, :]
    Y1_post = Y1[T0:T01]
    Y0_post = Y0[T0:T01, :]

    tau_mat = np.zeros(K)

    for k in range(K):
        # Holdout indices
        base = T0 - (r * K)
        start_idx = base + k * r
        end_idx = base + (k + 1) * r
        Hk = np.arange(start_idx, end_idx)

        # R: tau.mat[k,1] <- mean(Y1.post - rowMeans(Y0.post)) - mean(Y1.pre[Hk] - rowMeans(Y0.pre[Hk,]))
        tau_mat[k] = (np.mean(Y1_post - Y0_post.mean(axis=1)) - 
                      np.mean(Y1_pre[Hk] - Y0_pre[Hk, :].mean(axis=1)))

    tau_hat = np.mean(tau_mat)
    se_hat = np.sqrt(1 + ((K * r) / T1)) * np.std(tau_mat, ddof=1) / np.sqrt(K)
    t_hat = tau_hat / se_hat

    return {'t_hat': t_hat, 'tau_hat': tau_hat, 'se_hat': se_hat}
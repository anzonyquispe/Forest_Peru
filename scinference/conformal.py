import numpy as np
from estimators import did, sc, classo


def movingblock(Y1, Y0, T1, T0, theta0, estimation_method, lsei_type=1):
    """
    Moving block permutations for conformal inference.
    
    Equivalent to R's movingblock function.
    """
    T01 = T0 + T1
    
    # Handle theta0 (scalar or vector)
    theta0 = np.asarray(theta0).flatten()
    if len(theta0) == 1:
        theta0 = np.repeat(theta0, T1)

    # Adjust Y1 under the null: Y1_0 = Y1 - theta0 in post-treatment
    Y1_0 = Y1.copy()
    Y1_0[T0:T01] = Y1[T0:T01] - theta0

    # Get residuals
    if estimation_method == "classo":
        u_hat = classo(Y1_0, Y0)['u_hat']
    elif estimation_method == "sc":
        u_hat = sc(Y1_0, Y0, lsei_type)['u_hat']
    elif estimation_method == "did":
        u_hat = did(Y1_0, Y0)['u_hat']
    else:
        raise ValueError(f"Unknown estimation method: {estimation_method}")

    # Moving block test statistic
    sub_size = T1
    u_hat_c = np.concatenate([u_hat, u_hat])
    
    S_vec = np.zeros(T01)
    for s in range(T01):
        S_vec[s] = np.sum(np.abs(u_hat_c[s:s + sub_size]))

    # p-value
    p = np.mean(S_vec >= S_vec[T0])
    
    return p


def iid(Y1, Y0, T1, T0, theta0, estimation_method, n_perm, lsei_type=1):
    """
    IID permutations for conformal inference.
    
    Equivalent to R's iid function.
    """
    T01 = T0 + T1

    # Handle theta0 (scalar or vector)
    theta0 = np.asarray(theta0).flatten()
    if len(theta0) == 1:
        theta0 = np.repeat(theta0, T1)

    # Adjust Y1 under the null
    Y1_0 = Y1.copy()
    Y1_0[T0:T01] = Y1[T0:T01] - theta0

    # Get residuals
    if estimation_method == "classo":
        u_hat = classo(Y1_0, Y0)['u_hat']
    elif estimation_method == "sc":
        u_hat = sc(Y1_0, Y0, lsei_type)['u_hat']
    elif estimation_method == "did":
        u_hat = did(Y1_0, Y0)['u_hat']
    else:
        raise ValueError(f"Unknown estimation method: {estimation_method}")

    # Post-treatment indices
    post_ind = np.arange(T0, T01)
    
    # Observed test statistic
    Sq = np.sum(np.abs(u_hat[post_ind]))
    
    # Permutation distribution
    S_vec = np.zeros(n_perm)
    for r in range(n_perm):
        u_hat_p = np.random.permutation(u_hat)
        S_vec[r] = np.sum(np.abs(u_hat_p[post_ind]))

    # p-value
    p = (1 + np.sum(S_vec >= Sq)) / (n_perm + 1)
    
    return p


def confidence_interval(Y1, Y0, T1, T0, estimation_method, alpha, ci_grid, lsei_type=1):
    """
    Confidence interval via test inversion.
    
    Equivalent to R's confidence_interval function.
    """
    lb = np.full(T1, np.nan)
    ub = np.full(T1, np.nan)

    for t in range(T1):
        indices = list(range(T0)) + [T0 + t]
        Y1_temp = Y1[indices]
        Y0_temp = Y0[indices, :]

        ps_temp = np.zeros(len(ci_grid))

        for ind, theta in enumerate(ci_grid):
            Y1_0_temp = Y1_temp.copy()
            Y1_0_temp[T0] = Y1_temp[T0] - theta

            if estimation_method == "classo":
                u_hat = classo(Y1_0_temp, Y0_temp)['u_hat']
            elif estimation_method == "sc":
                u_hat = sc(Y1_0_temp, Y0_temp, lsei_type)['u_hat']
            elif estimation_method == "did":
                u_hat = did(Y1_0_temp, Y0_temp)['u_hat']
            else:
                raise ValueError(f"Unknown estimation method: {estimation_method}")

            ps_temp[ind] = np.mean(np.abs(u_hat) >= np.abs(u_hat[T0]))

        ci_temp = ci_grid[ps_temp > alpha]
        if len(ci_temp) > 0:
            lb[t] = np.min(ci_temp)
            ub[t] = np.max(ci_temp)

    return {'lb': lb, 'ub': ub}
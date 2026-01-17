import numpy as np
import cvxpy as cp


def did(Y1, Y0):
    """
    Difference-in-differences estimator.
    
    Equivalent to R:
        u.hat <- Y1 - mean(Y1-rowMeans(Y0)) - rowMeans(Y0)

    Parameters:
    -----------
    Y1 : numpy.ndarray
        Treated unit outcomes (T x 1 vector)
    Y0 : numpy.ndarray
        Control units outcomes (T x J matrix)

    Returns:
    --------
    dict : Dictionary containing:
        - u_hat: residuals (T x 1 vector)
    """
    Y1 = np.asarray(Y1).flatten()
    Y0 = np.asarray(Y0)

    Y0_row_means = Y0.mean(axis=1)
    u_hat = Y1 - np.mean(Y1 - Y0_row_means) - Y0_row_means

    return {'u_hat': u_hat}


def sc(Y1, Y0, lsei_type=1):
    """
    Synthetic control estimator.
    
    Solves:
        min ||Y0 @ w - Y1||^2
        subject to: sum(w) = 1, w >= 0
    
    Equivalent to R's limSolve::lsei

    Parameters:
    -----------
    Y1 : numpy.ndarray
        Treated unit outcomes (T x 1 vector)
    Y0 : numpy.ndarray
        Control units outcomes (T x J matrix)
    lsei_type : int, optional
        Type parameter (kept for compatibility, default=1)

    Returns:
    --------
    dict : Dictionary containing:
        - u_hat: residuals (T x 1 vector)
        - w_hat: weights (J x 1 vector)
    """
    Y1 = np.asarray(Y1).flatten()
    Y0 = np.asarray(Y0)

    J = Y0.shape[1]

    # Solve constrained least squares with cvxpy
    # min ||Y0 @ w - Y1||^2  s.t. sum(w) = 1, w >= 0
    w = cp.Variable(J)
    objective = cp.Minimize(cp.sum_squares(Y0 @ w - Y1))
    constraints = [cp.sum(w) == 1, w >= 0]

    problem = cp.Problem(objective, constraints)
    problem.solve(solver=cp.OSQP, verbose=False)

    if w.value is None:
        # Fallback solver if OSQP fails
        problem.solve(solver=cp.ECOS, verbose=False)

    w_hat = w.value
    u_hat = Y1 - Y0 @ w_hat

    return {'u_hat': u_hat, 'w_hat': w_hat}


def classo(Y1, Y0):
    """
    Constrained lasso estimator.
    
    Solves:
        min (1/T) * ||Y1 - [1, Y0] @ w||^2
        subject to: sum(|w[1:]|) <= 1
    
    Equivalent to R's CVXR implementation

    Parameters:
    -----------
    Y1 : numpy.ndarray
        Treated unit outcomes (T x 1 vector)
    Y0 : numpy.ndarray
        Control units outcomes (T x J matrix)

    Returns:
    --------
    dict : Dictionary containing:
        - u_hat: residuals (T x 1 vector)
        - w_hat: weights ((J+1) x 1 vector, including intercept)
    """
    Y1 = np.asarray(Y1).flatten()
    Y0 = np.asarray(Y0)

    T, J = Y0.shape

    # Add intercept column: X = [1, Y0]
    X = np.column_stack([np.ones(T), Y0])

    # min (1/T) * ||Y1 - X @ w||^2  s.t. sum(|w[1:]|) <= 1
    w = cp.Variable(J + 1)
    objective = cp.Minimize(cp.sum_squares(Y1 - X @ w) / T)
    constraints = [cp.norm1(w[1:]) <= 1]

    problem = cp.Problem(objective, constraints)
    problem.solve(solver=cp.OSQP, verbose=False)

    if w.value is None:
        problem.solve(solver=cp.ECOS, verbose=False)

    w_hat = w.value
    u_hat = Y1 - X @ w_hat

    return {'u_hat': u_hat, 'w_hat': w_hat}
"""
Python vs R Comparison Test
"""
import numpy as np
import pandas as pd
from scinference import scinference
# CSV path, test data R
path = r'C:\Users\Usuario\Documents\GitHub\scinference'
# TEST 1: CONFORMAL

print("=" * 50)
print("TEST CONFORMAL")
print("=" * 50)

Y0 = pd.read_csv(f'{path}/Y0_test.csv').values
Y1 = pd.read_csv(f'{path}/Y1_test.csv')['Y1'].values

T0 = 50
T1 = 5

# p-value
result = scinference(Y1, Y0, T1=T1, T0=T0, theta0=4, 
                     estimation_method="sc", permutation_method="mb")
print(f"\nP-value Python: {round(result['p_val'], 8)}")
print(f"P-value R:      0.01818182")

# confidence intervals
result_ci = scinference(Y1, Y0, T1=T1, T0=T0, estimation_method="sc", 
                        ci=True, ci_grid=np.arange(-2, 8.1, 0.1))
print(f"\nLB Python: {np.round(result_ci['lb'], 1)}")
print(f"LB R:      [ 0.8  1.5  0.4  1.8 -0.5]")
print(f"\nUB Python: {np.round(result_ci['ub'], 1)}")
print(f"UB R:      [4.1 4.6 4.  5.2 2.6]")

# TEST 2: T-TEST

print("\n" + "=" * 50)
print("TEST T-TEST")
print("=" * 50)

Y0 = pd.read_csv(f'{path}/Y0_ttest.csv').values
Y1 = pd.read_csv(f'{path}/Y1_ttest.csv')['Y1'].values

T0 = 30
T1 = 30

result = scinference(Y1, Y0, T1=T1, T0=T0, inference_method="ttest", K=2)
print(f"\nATT Python: {round(result['att'], 6)}")
print(f"ATT R:      1.488715")
print(f"\nSE Python:  {round(result['se'], 7)}")
print(f"SE R:       0.2922196")
print(f"\nLB Python:  {round(result['lb'], 6)}")
print(f"LB R:       -0.356287")
print(f"\nUB Python:  {round(result['ub'], 6)}")
print(f"UB R:       3.333716")

print("\n" + "=" * 50)
print("DONE")
print("=" * 50)
import pandas as pd
import matplotlib.pyplot as plt

ruta_excel = r"C:\era_data2\merge_partidos_final.xlsx"
df = pd.read_excel(ruta_excel, sheet_name="resultados")

if "RESUMEN" in df.columns:
    df = df[df["RESUMEN"].isna()]

df["gano_eleccion"] = (
    df.groupby("ubigeo")["total_votos"]
      .transform(lambda s: s == s.max())
      .astype(int)
)

# Usar todas las posiciones válidas (orden_aparicion >= 1)
mask_pos = df["orden_aparicion"].notna() & (df["orden_aparicion"] >= 1)

prob_pos = (
    df.loc[mask_pos]
      .groupby("orden_aparicion")
      .agg(
          n_ganan=("gano_eleccion", "sum"),
          N=("gano_eleccion", "size")
      )
      .reset_index()
)

prob_pos["prob_ganar"] = prob_pos["n_ganan"] / prob_pos["N"]
prob_pos = prob_pos.sort_values("orden_aparicion")

print("Resumen por posición en la boleta:")
for _, row in prob_pos.iterrows():
    pos = int(row["orden_aparicion"])
    n = int(row["N"])
    n_win = int(row["n_ganan"])
    p = row["prob_ganar"]
    print(f" - Posición {pos}: {n} partidos, {n_win} ganaron -> prob_ganar = {p:.3f}")

print("\nTabla completa:")
print(prob_pos)

# Gráfico
fig, ax = plt.subplots(figsize=(8, 5))
x_vals = prob_pos["orden_aparicion"].astype(str)
y_vals = prob_pos["prob_ganar"]

ax.bar(x_vals, y_vals)
for x, y in zip(x_vals, y_vals):
    ax.text(x, y + 0.01, f"{y:.3f}", ha="center", va="bottom", fontsize=9)

ax.set_xlabel("Posición en la boleta")
ax.set_ylabel("Probabilidad de ganar")
ax.set_title("Probabilidad de ganar por posición en la boleta")
ax.set_ylim(0, 1)

plt.tight_layout()
plt.show()

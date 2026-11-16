import openpyxl
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler

# Configuração de estilo visual
sns.set_theme(style="whitegrid")
plt.rcParams["font.sans-serif"] = "DejaVu Sans"

# 1. Carregar os dados originais
filepath = "ranking_pesquisadores_po_brasil_completo.xlsx"
df = pd.read_excel(filepath, sheet_name="Ranking")

# 2. Seleção de variáveis numéricas para o K-Means
features = [
    "Fator h",
    "Nº de artigos",
    "Citações totais",
    "Citações máx.",
    "Artigos de PO",
    "% em PO",
]

df_clean = df.dropna(subset=features).copy()


# Função para extrair o último sobrenome
def extrair_ultimo_sobrenome(nome):
    if pd.isna(nome):
        return ""
    partes = str(nome).strip().split()
    return partes[-1] if partes else ""


df_clean["Sobrenome"] = df_clean["Autor"].apply(extrair_ultimo_sobrenome)

X = df_clean[features]

# 3. Padronização das variáveis (z-score)
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X)

# 4. Agrupamento K-Means com k = 4
kmeans = KMeans(n_clusters=4, random_state=42, n_init=10)
df_clean["Cluster"] = kmeans.fit_predict(X_scaled)

# 5. PCA para projeção 2D
pca = PCA(n_components=2)
X_pca = pca.fit_transform(X_scaled)
df_clean["PCA1"] = X_pca[:, 0]
df_clean["PCA2"] = X_pca[:, 1]

# Reordenar colunas posicionando 'Cluster' e 'Sobrenome' em destaque
col_order = ["Ranking", "Autor", "Sobrenome", "Universidade", "Cluster"] + [
    c
    for c in df_clean.columns
    if c not in ["Ranking", "Autor", "Sobrenome", "Universidade", "Cluster"]
]
df_export = df_clean[col_order]

# 6. Salvar a planilha Excel com os clusters
output_excel = "ranking_pesquisadores_com_clusters.xlsx"
with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
    df_export.to_excel(
        writer, sheet_name="Pesquisadores com Clusters", index=False
    )
    resumo = df_clean.groupby("Cluster")[features].mean().round(2)
    resumo.insert(0, "Qtd Pesquisadores", df_clean["Cluster"].value_counts())
    resumo.to_excel(writer, sheet_name="Resumo dos Clusters")

print(f"Planilha exportada com sucesso: {output_excel}")

# ==============================================================================
# 7. CRIAR OS GRÁFICOS GERAIS (LIMPOS)
# ==============================================================================

# Figura 1: Scatter Plot PCA Geral
fig1, ax1 = plt.subplots(figsize=(10, 6))
sns.scatterplot(
    data=df_clean,
    x="PCA1",
    y="PCA2",
    hue="Cluster",
    style="Cluster",
    palette="tab10",
    s=100,
    alpha=0.85,
    ax=ax1,
)
ax1.set_title(
    "Clusters de Pesquisadores de PO (K-Means k=4 - Espaço PCA)",
    fontsize=14,
    fontweight="bold",
)
ax1.set_xlabel(
    f"Componente Principal 1 ({pca.explained_variance_ratio_[0]*100:.1f}% Variância Explicada)"
)
ax1.set_ylabel(
    f"Componente Principal 2 ({pca.explained_variance_ratio_[1]*100:.1f}% Variância Explicada)"
)
plt.tight_layout()
fig1.savefig("cluster_pca_kmeans.png", dpi=300)
plt.close(fig1)

# Figura 2: Dispersão Artigos vs. Citações Totais Geral
fig2, ax2 = plt.subplots(figsize=(10, 6))
sns.scatterplot(
    data=df_clean,
    x="Nº de artigos",
    y="Citações totais",
    hue="Cluster",
    size="Fator h",
    sizes=(40, 250),
    palette="tab10",
    alpha=0.8,
    ax=ax2,
)
ax2.set_title(
    "Nº de Artigos vs. Citações Totais por Cluster (Tamanho = Fator h)",
    fontsize=14,
    fontweight="bold",
)
ax2.set_xlabel("Nº de Artigos Totais")
ax2.set_ylabel("Citações Totais")
plt.tight_layout()
fig2.savefig("cluster_artigos_citacoes.png", dpi=300)
plt.close(fig2)

# Figura 3: Boxplots Comparativos
fig3, axes = plt.subplots(2, 3, figsize=(16, 9))
axes = axes.flatten()

for idx, col in enumerate(features):
    sns.boxplot(
        data=df_clean,
        x="Cluster",
        y=col,
        hue="Cluster",
        palette="tab10",
        ax=axes[idx],
    )
    axes[idx].set_title(f"Distribuição de {col}", fontsize=12)
    axes[idx].set_xlabel("Cluster")

plt.suptitle(
    "Comparação das Métricas entre os 4 Clusters", fontsize=16, fontweight="bold"
)
plt.tight_layout()
fig3.savefig("cluster_boxplots.png", dpi=300)
plt.close(fig3)

# ==============================================================================
# 8. CRIAR UM GRÁFICO INDIVIDUAL PARA CADA CLUSTER (COM SOBRENOMES)
# ==============================================================================

cores_palette = sns.color_palette("tab10", 4)

for cluster_id in sorted(df_clean["Cluster"].unique()):
    df_cluster = df_clean[df_clean["Cluster"] == cluster_id].copy()

    fig_c, ax_c = plt.subplots(figsize=(12, 8))

    # Scatter plot do cluster específico
    sns.scatterplot(
        data=df_cluster,
        x="Nº de artigos",
        y="Citações totais",
        size="Fator h",
        sizes=(60, 300),
        color=cores_palette[cluster_id],
        alpha=0.75,
        ax=ax_c,
    )

    # Offsets para evitar que o texto fique em cima do ponto
    x_max = df_cluster["Nº de artigos"].max()
    y_max = df_cluster["Citações totais"].max()
    x_offset = (x_max if x_max > 0 else 1) * 0.01
    y_offset = (y_max if y_max > 0 else 1) * 0.01

    # Adicionar o sobrenome de cada autor do cluster
    for _, row in df_cluster.iterrows():
        ax_c.text(
            row["Nº de artigos"] + x_offset,
            row["Citações totais"] + y_offset,
            row["Sobrenome"],
            fontsize=9,
            fontweight="semibold",
            alpha=0.9,
        )

    ax_c.set_title(
        f"Detalhamento do Cluster {cluster_id} (Total: {len(df_cluster)} pesquisadores)",
        fontsize=14,
        fontweight="bold",
    )
    ax_c.set_xlabel("Nº de Artigos Totais", fontsize=11)
    ax_c.set_ylabel("Citações Totais", fontsize=11)

    plt.tight_layout()
    filename = f"cluster_{cluster_id}_autores.png"
    fig_c.savefig(filename, dpi=300)
    plt.close(fig_c)
    print(f"Gráfico do Cluster {cluster_id} salvo como: {filename}")

print("\nTodos os arquivos foram gerados com sucesso!")
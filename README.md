# Software de Análise Estatística para Comparação de Grupos v1.01

Uma solução robusta e automatizada desenvolvida em Python para pesquisadores que necessitam comparar medidas entre grupos com rigor metodológico. O software automatiza a escolha dos testes estatísticos com base nos pressupostos dos dados, facilitando a análise de experimentos simples e complexos.

**Autor:** Eduardo Borba Neves

\---
<img width="1906" height="992" alt="image" src="https://github.com/user-attachments/assets/ff4d3673-e271-493d-a2ef-272f264ea55c" />

## 

## 🚀 Funcionalidades Principais

O App processa planilhas Excel (.xlsx) e realiza uma árvore de decisão estatística para aplicar o teste mais adequado ao seu desenho experimental:

### 

### 1\. Comparações Simples (Independentes)

* **Paramétrico:** Teste t de Student (ou Welch para variâncias desiguais) e ANOVA One-Way.
* **Não-Paramétrico:** Mann-Whitney U e Kruskal-Wallis.

### 

### 2\. Medidas Repetidas (Pareadas)

* **Paramétrico:** Teste t Pareado e ANOVA de Medidas Repetidas (RM-ANOVA).
* **Não-Paramétrico:** Wilcoxon Signed-Rank, Friedman e Teste de Quade.

### 

### 3\. Modelos com Covariáveis e Múltiplos Fatores

* **ANCOVA:** Análise de covariância para controle de variáveis interferentes.
* **ANCOVA Rank-Transformation:** Alternativa robusta para ANCOVA com dados não-normais.
* **ANOVA Fatorial:** Análise de dois fatores simultâneos.
* **Scheirer-Ray-Hare:** Extensão não-paramétrica para ANOVA de dois fatores.
* **Modelos Lineares Mistos (LME):** Utilizados como complemento para medidas repetidas.

\---

## 

## 📊 Rigor e Diagnóstico Estatístico

O software não apenas executa os testes, mas valida os pressupostos necessários:

* **Normalidade:** Testes de Shapiro-Wilk e Kolmogorov-Smirnov por grupo, acompanhados de gráficos Q-Q.
* **Homocedasticidade:** Teste de Levene para igualdade de variâncias.
* **Esfericidade:** Correção de Greenhouse-Geisser automática em RM-ANOVA (via pingouin).
* **Tamanho do Efeito:** Cálculo automático de d de Cohen, dz, r bisserial, Eta-quadrado ($\\eta^2$), Ômega-quadrado ($\\omega^2$) e R² Marginal.
* **Análise Post-Hoc:** Comparações múltiplas via Tukey HSD, Bonferroni, Nemenyi e Dunn.

\---

## 

## 🛠️ Requisitos e Instalação

A interface foi construída utilizando `customtkinter` para um visual moderno e profissional.

### 

### Dependências:

```bash
pip install customtkinter pandas numpy matplotlib seaborn scipy statsmodels openpyxl pingouin scikit-posthocs openpyxl
```

\---

## 💻 Como Utilizar

1. **Carregar Dados:** Importe sua planilha `.xlsx`.
2. **Mapear Variáveis:**

   * **Variável para Análise:** Sua variável dependente (numérica).
   * **Fator 1 (Grupo):** Sua variável independente principal.
   * **Fator 2 (Opcional):** Variável de tempo ou segundo grupo.
   * **ID (Opcional):** Necessário para medidas repetidas (vincula o sujeito aos momentos).
3. **Controle (Opcional):** Selecione variáveis de controle para realizar uma ANCOVA.
4. **Analisar:** O software gera um relatório técnico completo e quatro tipos de gráficos.
5. **Exportar:** Salve o relatório em `.txt` e as imagens em `.png` de alta resolução (300 DPI).

\---

## 

## 🎓 Citação Acadêmica

Se este software for utilizado em sua produção acadêmica, por favor, utilize a seguinte citação:

**Formato ABNT:**
NEVES, Eduardo Borba. **Software de análise estatística para comparação de medidas entre grupos**. Versão 1.01. \[S.l.]: Zenodo, 2026. Software. DOI: 10.5281/zenodo.19713014

**Formato APA:**
Neves, E. B. (2026). *Software de análise estatística para comparação de medidas entre grupos* (Version 1.01) \[Computer software]. Zenodo. https://doi.org/10.5281/zenodo.19713014

\---

## 

## 📄 Licença

Este projeto está licenciado sob a [Licença MIT](LICENSE).


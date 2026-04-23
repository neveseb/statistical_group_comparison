#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import customtkinter as ctk
from tkinter import filedialog, messagebox
import matplotlib.backends.backend_tkagg as tkagg
from scipy import stats
import statsmodels.api as sm
from statsmodels.formula.api import ols, mixedlm
import os

class StatisticalApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Software de análise estatística para comparação de medidas entre grupos v.1.01 - By Eduardo Borba Neves")
        self.geometry("1200x800")
        self.df = None
        self.checkboxes = {}
        self.figs = []
        self.setup_ui()
        self.protocol("WM_DELETE_WINDOW", self.safe_exit)

    def setup_ui(self):
        # 1. Sidebar com novo campo ID
        self.sidebar = ctk.CTkScrollableFrame(self, width=340, label_text="Configurações da Análise")
        self.sidebar.pack(side="left", fill="y", padx=10, pady=10)

        ctk.CTkButton(self.sidebar, text="Carregar Arquivo .XLSX para Começar", command=self.load_file, fg_color="#34495e").pack(pady=10, padx=10, fill="x")

        # Combox Grupo
        self.lbl_ev = ctk.CTkLabel(self.sidebar, text="Variável de Grupo ou Momento (Fator 1):", font=("Roboto", 12, "bold"))
        self.lbl_ev.pack(pady=(10,0)); self.cb_grupo = ctk.CTkComboBox(self.sidebar, command=self.refresh_checkboxes); self.cb_grupo.pack(pady=5)
        
        # Combox Variável Dependente
        self.lbl_gr = ctk.CTkLabel(self.sidebar, text="Variável para Análise:", font=("Roboto", 12, "bold"))
        self.lbl_gr.pack(pady=(10,0)); self.cb_varprincipal = ctk.CTkComboBox(self.sidebar, command=self.refresh_checkboxes); self.cb_varprincipal.pack(pady=5)

        # Combox Momentos
        self.lbl_te = ctk.CTkLabel(self.sidebar, text=" Opcional\n Variável de Grupo ou Momento (Fator 2):", font=("Roboto", 12, "bold"))
        self.lbl_te.pack(pady=(10,0)); self.cb_momentos = ctk.CTkComboBox(self.sidebar, command=self.refresh_checkboxes); self.cb_momentos.pack(pady=5)

        # NOVO: Combox ID para Medidas Repetidas
        self.lbl_id = ctk.CTkLabel(self.sidebar, text=" Opcional\n ID para Medidas Repetidas:", font=("Roboto", 12, "bold"))
        self.lbl_id.pack(pady=(10,0)); self.cb_id = ctk.CTkComboBox(self.sidebar, command=self.refresh_checkboxes); self.cb_id.pack(pady=5)
        
        # Checkboxes de Controle
        self.lbl_ctrl = ctk.CTkLabel(self.sidebar, text=" Opcional\n Fatores de Controle (ANCOVA):", font=("Roboto", 12, "bold"))
        self.lbl_ctrl.pack(pady=(20, 5))
        self.frame_checks = ctk.CTkFrame(self.sidebar)
        self.frame_checks.pack(fill="x", padx=5, pady=5)

        # Área Principal
        self.right_container = ctk.CTkFrame(self, fg_color="transparent")
        self.right_container.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        # Botões de Ação
        self.top_button_frame = ctk.CTkFrame(self.right_container)
        self.top_button_frame.pack(side="top", fill="x", pady=(0, 10))

        self.btn_run = ctk.CTkButton(self.top_button_frame, text="Rodar Análise Completa", fg_color="#27ae60", hover_color="#2ecc71", command=self.run_analysis)
        self.btn_run.pack(side="left", expand=True, fill="x", padx=5, pady=10)
        
        self.btn_export = ctk.CTkButton(self.top_button_frame, text="Exportar Resultados", fg_color="#f39c12", command=self.export_results)
        self.btn_export.pack(side="left", expand=True, fill="x", padx=5, pady=10)
        
        self.btn_exit = ctk.CTkButton(self.top_button_frame, text="Fechar e Sair", 
                                      fg_color="#c0392b", hover_color="#e74c3c", command=self.safe_exit)
        self.btn_exit.pack(side="left", expand=True, fill="x", padx=5, pady=10)

        # Painel de Resultados
        self.result_frame = ctk.CTkScrollableFrame(self.right_container, label_text="Relatório Técnico e Gráficos")
        self.result_frame.pack(side="bottom", fill="both", expand=True)
        
        self.text_result = ctk.CTkTextbox(self.result_frame, height=450, font=("Consolas", 11))
        self.text_result.pack(fill="x", pady=10)

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.df = pd.read_excel(path)
            cols = list(self.df.columns)
            for cb in [self.cb_grupo, self.cb_varprincipal, self.cb_momentos, self.cb_id]:
                cb.configure(values=[""] + cols)
                cb.set("")  # ← limpa a seleção atual
            self.refresh_checkboxes()
            messagebox.showinfo("Sucesso", "Dados carregados!")

    def refresh_checkboxes(self, _=None):
        if self.df is None: return
        for w in self.frame_checks.winfo_children(): w.destroy()
        self.checkboxes = {}
        selecionados = [self.cb_grupo.get(), self.cb_varprincipal.get(), self.cb_momentos.get(), self.cb_id.get()]
        for col in self.df.columns:
            if col not in selecionados and col != "":
                var = ctk.BooleanVar()
                ctk.CTkCheckBox(self.frame_checks, text=col, variable=var).pack(anchor="w", padx=10, pady=2)
                self.checkboxes[col] = var

    
    def run_analysis(self):
        if self.df is None:
            messagebox.showwarning("Aviso", "Por favor, carregue um arquivo Excel primeiro.")
            return

        # 1. LIMPEZA TOTAL DA ÁREA
        self.text_result.delete("1.0", "end")
        for w in self.result_frame.winfo_children():
            if not isinstance(w, ctk.CTkTextbox):
                w.destroy()
        for f in self.figs:
            plt.close(f)
        self.figs = []

        try:
            # ── 2. CAPTURA DAS VARIÁVEIS ──────────────────────────────────────────
            target     = str(self.cb_varprincipal.get()).strip()
            group      = str(self.cb_grupo.get()).strip()
            moment_raw = str(self.cb_momentos.get()).strip()
            id_raw     = str(self.cb_id.get()).strip()

            if not target or not group or target not in self.df.columns or group not in self.df.columns:
                messagebox.showerror("Erro de Seleção",
                                     "Selecione colunas válidas para Variável Principal e Grupo.")
                return

            moment_col = moment_raw if (moment_raw and moment_raw in self.df.columns) else None
            id_col     = id_raw    if (id_raw    and id_raw    in self.df.columns) else None
            controls   = [c for c, v in self.checkboxes.items() if v.get() and c in self.df.columns]

            # ── 3. FILTRAGEM DOS DADOS ────────────────────────────────────────────
            cols_to_use = [target, group]
            if moment_col: cols_to_use.append(moment_col)
            if id_col:     cols_to_use.append(id_col)
            cols_to_use.extend(controls)

            data = self.df[list(dict.fromkeys(cols_to_use))].dropna().copy()

            groups_list = data[group].unique()
            n_groups    = len(groups_list)
            samples     = [data[data[group] == g][target].values for g in groups_list]

            output = (
                "╔══════════════════════════════════════════════════════════════════════════╗\n"
                "║                       RELATÓRIO ESTATÍSTICO PRO                          ║\n"
                "╚══════════════════════════════════════════════════════════════════════════╝\n\n"
            )
            output += f"Variável dependente  : {target}\n"
            output += f"Fator principal      : {group}\n"
            if moment_col: output += f"Fator de momento     : {moment_col}\n"
            if id_col:     output += f"Identificador (ID)   : {id_col}\n"
            if controls:   output += f"Covariáveis          : {', '.join(controls)}\n"
            output += f"N total              : {len(data)}\n\n"

            # ── 4. ESTATÍSTICA DESCRITIVA ─────────────────────────────────────────
            output += "━"*75 + "\n"
            output += "1. ESTATÍSTICA DESCRITIVA\n"
            output += "━"*75 + "\n"

            grouping = [group]
            if moment_col: grouping.append(moment_col)

            desc = data.groupby(grouping)[target].agg(
                N='count', Média='mean', DP='std', Mediana='median',
                Mín='min', Máx='max', EPM='sem'
            )
            output += desc.to_string() + "\n\n"

            for g in groups_list:
                s = data[data[group] == g][target]
                output += (
                    f"  ► Grupo '{g}': N={len(s)}, "
                    f"M={s.mean():.2f} (DP={s.std():.2f}), "
                    f"Mediana={s.median():.2f}, "
                    f"[{s.min():.2f}–{s.max():.2f}]\n"
                )
            output += "\n"

            # ── 5. PRESSUPOSTOS ───────────────────────────────────────────────────


            # 5a. Normalidade (Shapiro-Wilk)
            output += "\n▸ Normalidade – Shapiro-Wilk:\n"
            is_normal = True
            sw_results = {}
            ks_results = {}
            for g in groups_list:
                sample = data[data[group] == g][target]
                if len(sample) >= 3:
                    stat_sw, p_sw = stats.shapiro(sample)
                    normal_sw = p_sw > 0.05
                    if not normal_sw: 
                        is_normal = False
                    sw_results[g] = (stat_sw, p_sw, normal_sw)
                    output += (
                        f"  Grupo '{g}': Shapiro-Wilk W={stat_sw:.4f}, p={p_sw:.4f} "
                        f"→ {'✔ Normal' if normal_sw else '✘ Não-Normal'}\n"
                    )

                    # Kolmogorov-Smirnov
                    # Para KS, precisamos comparar com uma distribuição normal com a média e DP da amostra
                    std_sample = sample.std(ddof=1)

                    if std_sample > 0:
                        stat_ks, p_ks = stats.kstest(sample, 'norm', args=(sample.mean(), std_sample))
                        normal_ks = p_ks > 0.05
                        if not normal_ks: 
                            is_normal = False
                        ks_results[g] = (stat_ks, p_ks, normal_ks)
                        output += (
                            f"  Grupo '{g}': Kolmogorov-Smirnov D={stat_ks:.4f}, p={p_ks:.4f} "
                            f"→ {'✔ Normal' if normal_ks else '✘ Não-Normal'}\n"
                        )
                    else:
                        ks_results[g] = (None, None, None)
                        output += (
                            f"  Grupo '{g}': Kolmogorov-Smirnov não aplicável (DP=0) → valores constantes\n"
                        )

                else:
                    output += f"  Grupo '{g}': N insuficiente para testes de normalidade (N={len(sample)}) → assumida normalidade\n"

            output += "\n"
            sw_interpretation = (
                f"  Interpretação: Os dados de '{target}' apresentam distribuição "
            )
            if is_normal:
                sw_interpretation += (
                    "normal em todos os grupos (p > 0,05 nos testes de Shapiro-Wilk e Kolmogorov-Smirnov), "
                    "justificando o uso de testes paramétricos.\n"
                )
            else:
                sw_interpretation += (
                    "não-normal em pelo menos um grupo (p ≤ 0,05 nos testes de Shapiro-Wilk ou Kolmogorov-Smirnov), "
                    "indicando o uso de testes não-paramétricos (exceto para modelos mistos, "
                    "que são robustos a desvios moderados da normalidade).\n"
                )
            output += sw_interpretation + "\n"

            # Levene — protegido: só faz sentido com ≥ 2 amostras não-vazias
            valid_samples = [s for s in samples if len(s) >= 2]
            if len(valid_samples) >= 2:
                stat_lev, p_lev = stats.levene(*valid_samples)
                homogeneous = p_lev > 0.05
                output += (
                    f"▸ Homogeneidade de Variâncias – Levene:\n"
                    f"  F={stat_lev:.4f}, p={p_lev:.4f} "
                    f"→ {'✔ Variâncias homogêneas' if homogeneous else '✘ Variâncias heterogêneas'}\n"
                    f"  Interpretação: As variâncias de '{target}' entre os grupos de '{group}' são "
                    f"{'homogêneas (pressuposto atendido).' if homogeneous else 'heterogêneas — será aplicada correção de Welch quando pertinente.'}\n\n"
                )
            else:
                homogeneous = True  # default seguro

            # ── 6. INFERÊNCIA E TAMANHO DO EFEITO ────────────────────────────────
            output += "━"*75 + "\n"
            output += "3. INFERÊNCIA ESTATÍSTICA E TAMANHO DO EFEITO\n"
            output += "━"*75 + "\n\n"

            posthoc_needed = False

            # ════════════════════════════════════════════════════════════════════
            # RAMO A – Medidas Repetidas (ID presente)
            # • 2 condições, normal    → t Pareado
            # • 2 condições, não-norm  → Wilcoxon pareado
            # • 3+ condições, normal   → RM-ANOVA (pingouin) + LME complementar
            # • 3+ condições, não-norm → Friedman + Quade + post-hoc Wilcoxon/Nemenyi
            # ════════════════════════════════════════════════════════════════════
            if id_col:

                # Coluna de repetição: moment_col preferido; senão group
                repeated_col = moment_col if moment_col else group

                # Pivot sujeito × condição — garante alinhamento independente da ordem
                pivot_all = data.pivot_table(
                    index=id_col, columns=repeated_col,
                    values=target, aggfunc='mean'
                ).dropna()   # mantém apenas sujeitos com TODAS as condições

                conditions          = list(pivot_all.columns)
                n_conditions        = len(conditions)
                n_subjects_complete = len(pivot_all)

                output += (
                    f"▸ Estrutura Detectada: Medidas Repetidas\n"
                    f"  Fator de repetição   : '{repeated_col}' ({n_conditions} condições/momentos)\n"
                    f"  Identificador (ID)   : '{id_col}'\n"
                    f"  Sujeitos com dados completos: {n_subjects_complete} "
                    f"(de {data[id_col].nunique()} únicos no dataset)\n\n"
                )

                # ── A1: DOIS MOMENTOS/CONDIÇÕES ──────────────────────────────────
                if n_conditions == 2:
                    cond_a, cond_b = conditions[0], conditions[1]

                    # Alinhamento correto via pivot — não depende da ordem das linhas
                    s_a    = pivot_all[cond_a].values
                    s_b    = pivot_all[cond_b].values
                    diffs  = s_a - s_b
                    n_pairs = len(pivot_all)

                    # Normalidade das diferenças
                    stat_sw_d, p_sw_d, diff_normal = np.nan, np.nan, True
                    if n_pairs >= 3:
                        stat_sw_d, p_sw_d = stats.shapiro(diffs)
                        diff_normal = p_sw_d > 0.05

                    if not np.isnan(stat_sw_d):
                        sw_d_txt = (
                            f"W={stat_sw_d:.4f}, p={p_sw_d:.4f} "
                            f"→ {'✔ Normal' if diff_normal else '✘ Não-Normal'}"
                        )
                    else:
                        sw_d_txt = (
                            f"N insuficiente para Shapiro-Wilk (N={n_pairs}) → assumida normalidade"
                        )
                    output += f"  Normalidade das diferenças (Shapiro-Wilk): {sw_d_txt}\n\n"

                    if diff_normal:
                        # ── t Pareado ────────────────────────────────────────────
                        res_tp    = stats.ttest_rel(s_a, s_b)
                        diff_mean = np.mean(diffs)
                        diff_sd   = np.std(diffs, ddof=1)
                        se_d      = diff_sd / np.sqrt(n_pairs)
                        t_crit_d  = stats.t.ppf(0.975, n_pairs - 1)
                        ic_lo_d   = diff_mean - t_crit_d * se_d
                        ic_hi_d   = diff_mean + t_crit_d * se_d
                        dz        = diff_mean / diff_sd if diff_sd > 0 else 0
                        interp_dz = "Grande" if abs(dz) >= 0.8 else "Médio" if abs(dz) >= 0.5 else "Pequeno"
                        sig_tp    = (
                            "estatisticamente significativa (p < 0,05)"
                            if res_tp.pvalue < 0.05 else "não significativa (p ≥ 0,05)"
                        )

                        output += (
                            f"▸ Modelo Selecionado: Teste t Pareado\n"
                            f"  Justificativa: 2 condições com medidas repetidas e diferenças "
                            f"com distribuição normal (SW p={p_sw_d:.4f}).\n\n"
                            f"  N de pares       : {n_pairs}\n"
                            f"  '{cond_a}': M={np.mean(s_a):.2f} (DP={np.std(s_a, ddof=1):.2f})\n"
                            f"  '{cond_b}': M={np.mean(s_b):.2f} (DP={np.std(s_b, ddof=1):.2f})\n"
                            f"  Diferença média  : {diff_mean:.3f} (DP={diff_sd:.3f})\n"
                            f"  IC 95% da dif.   : [{ic_lo_d:.3f}; {ic_hi_d:.3f}]\n"
                            f"  t({n_pairs-1}) = {res_tp.statistic:.3f}, p = {res_tp.pvalue:.4f}\n"
                            f"  Tamanho do efeito: dz de Cohen = {abs(dz):.3f} ({interp_dz})\n\n"
                            f"  ► Interpretação: A diferença em '{target}' entre os momentos "
                            f"'{cond_a}' (M={np.mean(s_a):.2f}) e '{cond_b}' (M={np.mean(s_b):.2f}) "
                            f"foi {sig_tp} para os sujeitos identificados por '{id_col}' "
                            f"[t({n_pairs-1})={res_tp.statistic:.3f}, p={res_tp.pvalue:.4f}]. "
                            f"O dz de Cohen de {abs(dz):.3f} indica efeito {interp_dz.lower()} "
                            f"(IC 95%: [{ic_lo_d:.3f}; {ic_hi_d:.3f}]).\n\n"
                        )

                    else:
                        # ── Wilcoxon Signed-Rank ─────────────────────────────────
                        res_wil = stats.wilcoxon(s_a, s_b, alternative='two-sided')
                        p_safe  = max(res_wil.pvalue, 1e-300)
                        z_wil   = abs(stats.norm.ppf(p_safe / 2))
                        r_wil   = min(z_wil / np.sqrt(n_pairs), 1.0)
                        interp_r_wil = "Grande" if r_wil >= 0.5 else "Médio" if r_wil >= 0.3 else "Pequeno"
                        sig_wil = (
                            "estatisticamente significativa (p < 0,05)"
                            if res_wil.pvalue < 0.05 else "não significativa (p ≥ 0,05)"
                        )

                        output += (
                            f"▸ Modelo Selecionado: Wilcoxon Signed-Rank (não-paramétrico, pareado)\n"
                            f"  Justificativa: 2 condições com medidas repetidas e diferenças "
                            f"com distribuição não-normal (SW p={p_sw_d:.4f}).\n\n"
                            f"  N de pares        : {n_pairs}\n"
                            f"  '{cond_a}': Md={np.median(s_a):.2f}\n"
                            f"  '{cond_b}': Md={np.median(s_b):.2f}\n"
                            f"  W = {res_wil.statistic:.1f}, p = {res_wil.pvalue:.4f}\n"
                            f"  Tamanho do efeito : r = {r_wil:.3f} ({interp_r_wil})\n\n"
                            f"  ► Interpretação: A diferença em '{target}' entre os momentos "
                            f"'{cond_a}' (Md={np.median(s_a):.2f}) e '{cond_b}' (Md={np.median(s_b):.2f}) "
                            f"foi {sig_wil} para os sujeitos identificados por '{id_col}' "
                            f"[Wilcoxon: W={res_wil.statistic:.1f}, p={res_wil.pvalue:.4f}]. "
                            f"O r de {r_wil:.3f} indica efeito {interp_r_wil.lower()} "
                            f"da condição sobre '{target}'.\n\n"
                        )

                # ── A2: TRÊS OU MAIS MOMENTOS/CONDIÇÕES ──────────────────────────
                else:
                    mat         = pivot_all.values          # (n_subjects × n_conditions)
                    n_sub       = mat.shape[0]
                    cond_labels = list(pivot_all.columns)

                    if is_normal:
                        # ── RM-ANOVA via pingouin + LME complementar ─────────────
                        output += (
                            f"▸ Modelo Selecionado: ANOVA de Medidas Repetidas (RM-ANOVA)\n"
                            f"  Justificativa: {n_conditions} condições/momentos, medidas repetidas "
                            f"(ID='{id_col}'), distribuição normal. A RM-ANOVA remove a variabilidade "
                            f"individual do erro, aumentando a sensibilidade.\n\n"
                        )

                        try:
                            import pingouin as pg

                            rm = pg.rm_anova(
                                data=data, dv=target,
                                within=repeated_col, subject=id_col,
                                correction=True, detailed=True
                            )
                            output += "  Tabela RM-ANOVA (com correção de esfericidade de Mauchly):\n"
                            output += rm.to_string() + "\n\n"

                            def _get(col):
                                v = rm.loc[rm['Source'] == repeated_col, col].values
                                return v[0] if len(v) > 0 else np.nan

                            p_rm   = _get('p-unc')
                            f_rm   = _get('F')
                            ng2_rm = _get('ng2')
                            eps_rm = _get('eps')

                            # limiares η²g de Bakeman (2005)
                            interp_ng2 = (
                                "Grande" if (not np.isnan(ng2_rm) and ng2_rm >= 0.26) else
                                "Médio"  if (not np.isnan(ng2_rm) and ng2_rm >= 0.13) else "Pequeno"
                            )
                            sig_rm = (
                                "estatisticamente significativo (p < 0,05)"
                                if (not np.isnan(p_rm) and p_rm < 0.05) else "não significativo (p ≥ 0,05)"
                            )

                            sph_line = ""
                            if not np.isnan(eps_rm):
                                sph_ok   = eps_rm >= 0.75
                                sph_line = (
                                    f"\n  Esfericidade (ε de Greenhouse-Geisser) = {eps_rm:.3f} "
                                    f"→ {'✔ Aceitável' if sph_ok else '✘ Violada — correção GG aplicada'}"
                                )

                            output += (
                                f"  F = {f_rm:.3f}, p = {p_rm:.4f}, η²g = {ng2_rm:.3f} ({interp_ng2})"
                                f"{sph_line}\n\n"
                                f"  ► Interpretação: A RM-ANOVA revelou que o efeito de '{repeated_col}' "
                                f"sobre '{target}' foi {sig_rm} [F={f_rm:.3f}, p={p_rm:.4f}]. "
                                f"O η²g de {ng2_rm:.3f} indica efeito {interp_ng2.lower()}. "
                            )
                            if not np.isnan(eps_rm) and eps_rm < 0.75:
                                output += (
                                    "A esfericidade foi violada (ε < 0,75); a correção de "
                                    "Greenhouse-Geisser foi aplicada automaticamente. "
                                )
                            output += "\n\n"

                            # LME complementar
                            output += "  Complemento — Modelo Linear Misto (LME):\n"
                            formula_lme = f"Q('{target}') ~ C(Q('{repeated_col}'))"
                            for c in controls:
                                formula_lme += f" + Q('{c}')"
                            model_lme = mixedlm(formula_lme, data, groups=data[id_col]).fit(reml=True)
                            output += model_lme.summary().as_text() + "\n\n"

                            # R² Marginal — Nakagawa & Schielzeth (2013)
                            var_fixed   = np.var(model_lme.fittedvalues - model_lme.resid)
                            var_resid   = np.var(model_lme.resid)
                            re_var      = float(model_lme.cov_re.values[0, 0]) if model_lme.cov_re is not None else 0.0
                            r2_marg     = max(min(var_fixed / (var_fixed + re_var + var_resid), 1.0), 0.0)
                            interp_r2   = "Grande" if r2_marg > 0.26 else "Médio" if r2_marg > 0.13 else "Pequeno"
                            output += f"  R² Marginal (LME, Nakagawa & Schielzeth 2013) = {r2_marg:.3f} ({interp_r2})\n\n"

                            # Post-hoc Bonferroni se RM-ANOVA significativa e ≥ 3 condições
                            if not np.isnan(p_rm) and p_rm < 0.05:
                                output += "━"*75 + "\n"
                                output += "4. ANÁLISE POST-HOC – Bonferroni Pareado (RM-ANOVA)\n"
                                output += "━"*75 + "\n\n"

                                ph_rm = pg.pairwise_tests(
                                    data=data, dv=target,
                                    within=repeated_col, subject=id_col,
                                    padjust='bonf', parametric=True
                                )
                                cols_ph = [c for c in ['A','B','T','dof','p-unc','p-corr','p-adjust','cohen-d','BF10']
                                           if c in ph_rm.columns]
                                output += "▸ Comparações pareadas – Bonferroni:\n"
                                output += ph_rm[cols_ph].to_string(index=False) + "\n\n"
                                output += (
                                    f"  ► Interpretação: Pares com p-corr < 0,05 diferem "
                                    f"significativamente em '{target}'. O d de Cohen indica a "
                                    f"magnitude da diferença por par.\n\n"
                                )

                        except ImportError:
                            # Fallback: LME apenas
                            output += (
                                "  (pingouin não instalado — usando LME como fallback.\n"
                                "  Instale com: pip install pingouin)\n\n"
                            )
                            formula_lme = f"Q('{target}') ~ C(Q('{repeated_col}'))"
                            for c in controls:
                                formula_lme += f" + Q('{c}')"
                            model_lme = mixedlm(formula_lme, data, groups=data[id_col]).fit(reml=True)
                            output += "  Sumário LME:\n"
                            output += model_lme.summary().as_text() + "\n\n"

                            var_fixed = np.var(model_lme.fittedvalues - model_lme.resid)
                            var_resid = np.var(model_lme.resid)
                            re_var    = float(model_lme.cov_re.values[0, 0]) if model_lme.cov_re is not None else 0.0
                            r2_marg   = max(min(var_fixed / (var_fixed + re_var + var_resid), 1.0), 0.0)
                            interp_r2 = "Grande" if r2_marg > 0.26 else "Médio" if r2_marg > 0.13 else "Pequeno"
                            output += f"  R² Marginal = {r2_marg:.3f} ({interp_r2})\n\n"

                            # Extrair p do efeito principal
                            fe_pvalues = model_lme.pvalues
                            rc_terms   = [k for k in fe_pvalues.index if repeated_col in str(k)]
                            min_p_lme  = fe_pvalues[rc_terms].min() if rc_terms else np.nan
                            sig_lme    = (
                                "significativo (p < 0,05)"
                                if (not np.isnan(min_p_lme) and min_p_lme < 0.05)
                                else "não significativo (p ≥ 0,05)"
                            )
                            output += (
                                f"  ► O LME revelou que o efeito de '{repeated_col}' "
                                f"sobre '{target}' foi {sig_lme}.\n\n"
                            )

                            # Post-hoc: t pareado pairwise + Bonferroni (via pivot_all)
                            if not np.isnan(min_p_lme) and min_p_lme < 0.05:
                                output += "━"*75 + "\n"
                                output += "4. ANÁLISE POST-HOC – t Pareado + Bonferroni (LME fallback)\n"
                                output += "━"*75 + "\n\n"
                                n_comp_lme = n_conditions * (n_conditions - 1) // 2
                                output += "▸ Comparações pareadas (t pareado + Bonferroni):\n"
                                for i_c, c1 in enumerate(cond_labels):
                                    for c2 in cond_labels[i_c+1:]:
                                        col1 = pivot_all[c1].values
                                        col2 = pivot_all[c2].values
                                        if len(col1) < 2:
                                            continue
                                        res_ph   = stats.ttest_rel(col1, col2)
                                        p_adj_ph = min(res_ph.pvalue * n_comp_lme, 1.0)
                                        diff_ph  = np.mean(col1 - col2)
                                        sd_ph    = np.std(col1 - col2, ddof=1)
                                        dz_ph    = abs(diff_ph) / sd_ph if sd_ph > 0 else 0
                                        interp_dz_ph = "Grande" if dz_ph >= 0.8 else "Médio" if dz_ph >= 0.5 else "Pequeno"
                                        sig_ph = "✔" if p_adj_ph < 0.05 else "✘"
                                        output += (
                                            f"  {sig_ph} '{c1}' vs '{c2}': "
                                            f"t={res_ph.statistic:.3f}, p_ajust={p_adj_ph:.4f}, "
                                            f"dz={dz_ph:.3f} ({interp_dz_ph})\n"
                                        )
                                output += "\n"

                    else:
                        # ── Friedman + Quade (não-paramétrico, 3+ condições repetidas) ──
                        output += (
                            f"▸ Modelo Selecionado: Friedman + Quade (não-paramétrico, medidas repetidas)\n"
                            f"  Justificativa: {n_conditions} condições/momentos com medidas repetidas "
                            f"(ID='{id_col}') e distribuição não-normal. Friedman e Quade são as "
                            f"alternativas não-paramétricas à RM-ANOVA.\n\n"
                        )

                        # ── Friedman ──────────────────────────────────────────────
                        res_fr  = stats.friedmanchisquare(*[mat[:, j] for j in range(n_conditions)])
                        W_fr    = res_fr.statistic / (n_sub * (n_conditions - 1))
                        interp_W = "Grande" if W_fr >= 0.7 else "Médio" if W_fr >= 0.5 else "Pequeno"
                        sig_fr  = (
                            "estatisticamente significativo (p < 0,05)"
                            if res_fr.pvalue < 0.05 else "não significativo (p ≥ 0,05)"
                        )
                        output += (
                            f"  ── Friedman:\n"
                            f"  χ²({n_conditions-1}) = {res_fr.statistic:.3f}, p = {res_fr.pvalue:.4f}\n"
                            f"  W de Kendall = {W_fr:.3f} ({interp_W})\n"
                            f"  ► O teste de Friedman revelou que o efeito de '{repeated_col}' "
                            f"sobre '{target}' foi {sig_fr} "
                            f"[χ²({n_conditions-1})={res_fr.statistic:.3f}, p={res_fr.pvalue:.4f}]. "
                            f"O W de Kendall de {W_fr:.3f} indica concordância {interp_W.lower()} "
                            f"entre os postos das condições.\n\n"
                        )

                        # ── Quade (Conover, 1999) ─────────────────────────────────
                        ranks_within = np.apply_along_axis(stats.rankdata, 1, mat)
                        ranges_sub   = mat.max(axis=1) - mat.min(axis=1)
                        ranges_rank  = stats.rankdata(ranges_sub)
                        S            = ranges_rank[:, np.newaxis] * (ranks_within - (n_conditions + 1) / 2)
                        S_col        = S.sum(axis=0)
                        A1           = (S**2).sum()
                        SS_treat_q   = (S_col**2).sum() / n_sub
                        SS_error_q   = A1 - SS_treat_q
                        df1_q        = n_conditions - 1
                        df2_q        = (n_sub - 1) * (n_conditions - 1)

                        if SS_error_q > 0:
                            F_q = ((n_sub - 1) * SS_treat_q) / SS_error_q
                            p_q = 1 - stats.f.cdf(F_q, df1_q, df2_q)
                        else:
                            F_q, p_q = np.nan, np.nan

                        sig_q = (
                            "estatisticamente significativo (p < 0,05)"
                            if (not np.isnan(p_q) and p_q < 0.05) else
                            "não significativo (p ≥ 0,05)" if not np.isnan(p_q) else "não calculável"
                        )
                        output += (
                            f"  ── Quade:\n"
                            f"  F({df1_q},{df2_q}) = {F_q:.3f}, p = {p_q:.4f}\n"
                            f"  ► O teste de Quade revelou que o efeito de '{repeated_col}' "
                            f"sobre '{target}' foi {sig_q} "
                            f"[F({df1_q},{df2_q})={F_q:.3f}, p={p_q:.4f}].\n\n"
                        )

                        # ── Post-hoc: Nemenyi + Wilcoxon pareado + Bonferroni ─────
                        p_fr_sig = res_fr.pvalue < 0.05
                        p_q_sig  = (not np.isnan(p_q)) and p_q < 0.05
                        if p_fr_sig or p_q_sig:
                            output += "━"*75 + "\n"
                            output += "4. ANÁLISE POST-HOC – Wilcoxon Pareado + Bonferroni\n"
                            output += "━"*75 + "\n\n"

                            # Nemenyi (scikit-posthocs, opcional)
                            try:
                                from scikit_posthocs import posthoc_nemenyi_friedman
                                nem         = posthoc_nemenyi_friedman(mat)
                                nem.index   = cond_labels
                                nem.columns = cond_labels
                                output += "▸ Nemenyi post-hoc (Friedman):\n"
                                output += nem.to_string() + "\n\n"
                                output += (
                                    f"  ► Valores p < 0,05 indicam diferença significativa entre o par "
                                    f"em '{target}'.\n\n"
                                )
                            except ImportError:
                                pass   # Wilcoxon abaixo cobre o caso

                            # Wilcoxon pareado pairwise + Bonferroni (sempre executado)
                            n_comp_fr = n_conditions * (n_conditions - 1) // 2
                            output += "▸ Wilcoxon Signed-Rank pareado + Bonferroni:\n"
                            for i_c in range(n_conditions):
                                for j_c in range(i_c + 1, n_conditions):
                                    c1_l = cond_labels[i_c]
                                    c2_l = cond_labels[j_c]
                                    col1 = mat[:, i_c]
                                    col2 = mat[:, j_c]
                                    if len(col1) < 2:
                                        continue
                                    res_wil_ph  = stats.wilcoxon(col1, col2, alternative='two-sided')
                                    p_adj_wil   = min(res_wil_ph.pvalue * n_comp_fr, 1.0)
                                    p_safe_ph   = max(res_wil_ph.pvalue, 1e-300)
                                    z_wil_ph    = abs(stats.norm.ppf(p_safe_ph / 2))
                                    r_wil_ph    = min(z_wil_ph / np.sqrt(n_sub), 1.0)
                                    interp_r_ph = "Grande" if r_wil_ph >= 0.5 else "Médio" if r_wil_ph >= 0.3 else "Pequeno"
                                    sig_wil_ph  = "✔" if p_adj_wil < 0.05 else "✘"
                                    output += (
                                        f"  {sig_wil_ph} '{c1_l}' vs '{c2_l}': "
                                        f"W={res_wil_ph.statistic:.1f}, p_ajust={p_adj_wil:.4f}, "
                                        f"r={r_wil_ph:.3f} ({interp_r_ph})\n"
                                    )
                            output += "\n"

            # ════════════════════════════════════════════════════════════════════
            # RAMO B – 2 Fatores / ANCOVA (sem ID, mas com momento ou covariáveis)
            # Normal   → ANOVA Fatorial / ANCOVA  (post-hoc: Tukey HSD)
            # Não-norm + controls → ANCOVA Rank-Transform (post-hoc: Tukey sobre postos)
            # Não-norm + moment   → Scheirer-Ray-Hare  (post-hoc: Dunn + Bonferroni)
            # ════════════════════════════════════════════════════════════════════
            elif moment_col or controls:

                if is_normal:
                    # ── B1: ANOVA Fatorial ou ANCOVA ─────────────────────────────
                    model_name = "ANCOVA" if controls else "ANOVA Fatorial (2 Fatores)"
                    output += f"▸ Modelo Selecionado: {model_name}\n"
                    output += (
                        f"  Justificativa: "
                        f"{'Covariáveis (' + ', '.join(controls) + ') presentes — ANCOVA.' if controls else 'Dois fatores categóricos (' + group + ' e ' + moment_col + ') com distribuição normal — ANOVA fatorial.'}\n\n"
                    )

                    formula = f"Q('{target}') ~ C(Q('{group}'))"
                    if moment_col:
                        formula += f" * C(Q('{moment_col}'))"
                    for c in controls:
                        formula += f" + Q('{c}')"

                    res_ols = ols(formula, data).fit()
                    table   = sm.stats.anova_lm(res_ols, typ=2)
                    output += "  Tabela ANOVA (Tipo II – Soma dos Quadrados):\n"
                    output += table.to_string() + "\n\n"

                    output += "  Tamanho do Efeito – Eta-quadrado Parcial (η²p):\n"
                    ss_res = table.loc['Residual', 'sum_sq']
                    p_group_b1 = np.nan

                    for factor in table.index:
                        if factor == 'Residual':
                            continue
                        ss_f   = table.loc[factor, 'sum_sq']
                        f_val  = table.loc[factor, 'F']
                        p_val  = table.loc[factor, 'PR(>F)']
                        eta_p  = ss_f / (ss_f + ss_res)
                        interp_eta = "Grande" if eta_p >= 0.14 else "Médio" if eta_p >= 0.06 else "Pequeno"
                        sig    = "p < 0,05 ✔" if p_val < 0.05 else "p ≥ 0,05 ✘"
                        output += (
                            f"  • {factor}: F={f_val:.3f}, p={p_val:.4f} ({sig}), "
                            f"η²p={eta_p:.3f} ({interp_eta})\n"
                        )
                        if group in str(factor) and ':' not in str(factor):
                            p_group_b1 = p_val

                    sig_b1 = (
                        "estatisticamente significativo (p < 0,05)"
                        if (not np.isnan(p_group_b1) and p_group_b1 < 0.05) else "não significativo (p ≥ 0,05)"
                    )
                    output += f"\n  ► Interpretação: A {model_name} revelou que o efeito principal de '{group}' sobre '{target}' foi {sig_b1}. "

                    if controls:
                        output += f"Após controle das covariáveis ({', '.join(controls)}), o efeito de '{group}' é avaliado de forma mais precisa. "

                    if moment_col:
                        p_inter_row = [i for i in table.index if ':' in str(i)]
                        if p_inter_row:
                            p_inter = table.loc[p_inter_row[0], 'PR(>F)']
                            output += (
                                f"A interação '{group} × {moment_col}' foi "
                                + (f"significativa (p={p_inter:.4f}), sugerindo que o efeito de '{group}' "
                                   f"sobre '{target}' varia conforme o momento avaliado."
                                   if p_inter < 0.05 else
                                   f"não significativa (p={p_inter:.4f}), indicando efeitos independentes dos fatores.")
                            )
                    output += "\n\n"

                    # Post-hoc Tukey HSD (grupo principal sig. e > 2 grupos)
                    if not np.isnan(p_group_b1) and p_group_b1 < 0.05 and n_groups > 2:
                        output += "━"*75 + "\n"
                        output += "4. ANÁLISE POST-HOC – Tukey HSD (ANOVA/ANCOVA)\n"
                        output += "━"*75 + "\n\n"
                        from statsmodels.stats.multicomp import pairwise_tukeyhsd
                        tukey = pairwise_tukeyhsd(data[target], data[group], alpha=0.05)
                        output += "▸ Tukey HSD:\n"
                        output += str(tukey.summary()) + "\n\n"
                        output += "  Tamanho do Efeito pairwise – d de Cohen:\n"
                        for i_g, g1 in enumerate(groups_list):
                            for g2 in groups_list[i_g+1:]:
                                s1_ph = data[data[group] == g1][target].values
                                s2_ph = data[data[group] == g2][target].values
                                sp_ph = np.sqrt(
                                    ((len(s1_ph)-1)*np.std(s1_ph, ddof=1)**2 +
                                     (len(s2_ph)-1)*np.std(s2_ph, ddof=1)**2) /
                                    (len(s1_ph)+len(s2_ph)-2)
                                )
                                d_ph = abs(np.mean(s1_ph) - np.mean(s2_ph)) / sp_ph if sp_ph > 0 else 0
                                interp_d_ph = "Grande" if d_ph >= 0.8 else "Médio" if d_ph >= 0.5 else "Pequeno"
                                output += f"  • '{g1}' vs '{g2}': d={d_ph:.3f} ({interp_d_ph})\n"
                        output += (
                            f"\n  ► Pares com p ajustado < 0,05 diferem significativamente "
                            f"em '{target}'.\n\n"
                        )

                else:
                    # ── B2: NÃO-PARAMÉTRICO ───────────────────────────────────────
                    if controls:
                        # B2a: ANCOVA Rank-Transformation
                        output += "▸ Modelo Selecionado: ANCOVA Baseada em Postos (Rank-Transformation)\n"
                        output += (
                            f"  Justificativa: Dados não-normais com covariáveis. A ANCOVA baseada em "
                            f"postos transforma os dados em ranks e aplica ANCOVA paramétrica, "
                            f"sendo robusta a desvios da normalidade.\n\n"
                        )

                        data_rank = data.copy()
                        data_rank[f'{target}_rank'] = data_rank[target].rank()
                        for c in controls:
                            data_rank[f'{c}_rank'] = data_rank[c].rank()

                        formula_rank = f"Q('{target}_rank') ~ C(Q('{group}'))"
                        if moment_col:
                            formula_rank += f" * C(Q('{moment_col}'))"
                        for c in controls:
                            formula_rank += f" + Q('{c}_rank')"

                        res_ols_rank = ols(formula_rank, data_rank).fit()
                        table_rank   = sm.stats.anova_lm(res_ols_rank, typ=2)
                        output += "  Tabela ANCOVA sobre Postos (Tipo II):\n"
                        output += table_rank.to_string() + "\n\n"

                        output += "  Tamanho do Efeito – η²p sobre Postos:\n"
                        ss_res_rank = table_rank.loc['Residual', 'sum_sq']
                        p_group_b2a = np.nan

                        for factor in table_rank.index:
                            if factor == 'Residual':
                                continue
                            ss_f_r  = table_rank.loc[factor, 'sum_sq']
                            f_val_r = table_rank.loc[factor, 'F']
                            p_val_r = table_rank.loc[factor, 'PR(>F)']
                            eta_p_r = ss_f_r / (ss_f_r + ss_res_rank)
                            interp_r = "Grande" if eta_p_r >= 0.14 else "Médio" if eta_p_r >= 0.06 else "Pequeno"
                            sig_r   = "p < 0,05 ✔" if p_val_r < 0.05 else "p ≥ 0,05 ✘"
                            output += (
                                f"  • {factor}: F={f_val_r:.3f}, p={p_val_r:.4f} ({sig_r}), "
                                f"η²p={eta_p_r:.3f} ({interp_r})\n"
                            )
                            if group in str(factor) and ':' not in str(factor):
                                p_group_b2a = p_val_r

                        sig_b2a = (
                            "estatisticamente significativo (p < 0,05)"
                            if (not np.isnan(p_group_b2a) and p_group_b2a < 0.05) else "não significativo (p ≥ 0,05)"
                        )
                        output += (
                            f"\n  ► A ANCOVA baseada em postos revelou que o efeito de '{group}' "
                            f"sobre os postos de '{target}' foi {sig_b2a}. "
                            f"Após controle das covariáveis rank-transformadas ({', '.join(controls)}), "
                            f"o efeito de '{group}' é avaliado de forma mais precisa.\n\n"
                        )

                        # Post-hoc Tukey sobre postos
                        if not np.isnan(p_group_b2a) and p_group_b2a < 0.05 and n_groups > 2:
                            output += "━"*75 + "\n"
                            output += "4. ANÁLISE POST-HOC – Tukey HSD sobre Postos (ANCOVA Rank-Transform)\n"
                            output += "━"*75 + "\n\n"
                            from statsmodels.stats.multicomp import pairwise_tukeyhsd
                            tukey_rank = pairwise_tukeyhsd(
                                data_rank[f'{target}_rank'], data_rank[group], alpha=0.05
                            )
                            output += "▸ Tukey HSD sobre Postos:\n"
                            output += str(tukey_rank.summary()) + "\n\n"
                            output += (
                                f"  ► Pares com p ajustado < 0,05 diferem significativamente "
                                f"nos postos de '{target}'.\n\n"
                            )

                    else:
                        # B2b: Scheirer-Ray-Hare (2 fatores sem covariáveis)
                        output += "▸ Modelo Selecionado: Scheirer-Ray-Hare (não-paramétrico, 2 fatores)\n"
                        output += (
                            f"  Justificativa: Dados não-normais com dois fatores categóricos "
                            f"('{group}' e '{moment_col}'). Scheirer-Ray-Hare é a extensão não-paramétrica "
                            f"da ANOVA de dois fatores, operando sobre postos.\n\n"
                        )

                        data_srh = data.copy()
                        data_srh['_rank_'] = data_srh[target].rank()
                        N_srh = len(data_srh)
                        grand_mean_rank = data_srh['_rank_'].mean()
                        SS_total = ((data_srh['_rank_'] - grand_mean_rank)**2).sum()

                        def ss_factor(df, rank_col, factor_col):
                            grp_means = df.groupby(factor_col)[rank_col].mean()
                            grp_ns    = df.groupby(factor_col)[rank_col].count()
                            return sum(
                                grp_ns[g] * (grp_means[g] - df[rank_col].mean())**2
                                for g in grp_means.index
                            )

                        srh_factors = [group]
                        if moment_col:
                            srh_factors.append(moment_col)

                        srh_results = {}
                        for fac in srh_factors:
                            df_fac = len(data_srh[fac].unique()) - 1
                            ss_fac = ss_factor(data_srh, '_rank_', fac)
                            ms_fac = ss_fac / df_fac if df_fac > 0 else np.nan
                            H_fac  = ms_fac / (SS_total / (N_srh - 1)) if ms_fac else np.nan
                            p_fac  = 1 - stats.chi2.cdf(H_fac, df_fac) if not np.isnan(H_fac) else np.nan
                            srh_results[fac] = {
                                'df': df_fac, 'SS': ss_fac,
                                'MS': ms_fac, 'H': H_fac, 'p': p_fac
                            }

                        # Interação via SS das células
                        srh_inter_p = np.nan
                        if moment_col:
                            data_srh['_cell_'] = (
                                data_srh[group].astype(str) + '_' + data_srh[moment_col].astype(str)
                            )
                            ss_cells  = ss_factor(data_srh, '_rank_', '_cell_')
                            ss_inter  = ss_cells - srh_results[group]['SS'] - srh_results[moment_col]['SS']
                            df_inter  = srh_results[group]['df'] * srh_results[moment_col]['df']
                            ms_inter  = ss_inter / df_inter if df_inter > 0 else np.nan
                            H_inter   = ms_inter / (SS_total / (N_srh - 1)) if ms_inter else np.nan
                            p_inter_s = 1 - stats.chi2.cdf(H_inter, df_inter) if not np.isnan(H_inter) else np.nan
                            srh_results['Interação'] = {
                                'df': df_inter, 'SS': ss_inter,
                                'MS': ms_inter, 'H': H_inter, 'p': p_inter_s
                            }
                            srh_inter_p = p_inter_s

                        output += "  Tabela Scheirer-Ray-Hare (baseada em postos):\n"
                        output += (
                            f"  {'Fator':<30} {'df':>4} {'SS':>10} {'H':>8} "
                            f"{'p':>8} {'Sig':>4} {'η²H':>7} {'Magnitude':>10}\n"
                        )
                        output += "  " + "-"*75 + "\n"

                        srh_group_p = np.nan
                        srh_group_H = np.nan
                        srh_group_df = 1
                        srh_group_eta = 0.0
                        srh_group_interp = "Pequeno"

                        for fac, res_srh in srh_results.items():
                            if np.isnan(res_srh.get('p', np.nan)):
                                continue
                            eta2H      = max(res_srh['H'] / (N_srh - 1), 0)
                            interp_srh = "Grande" if eta2H >= 0.14 else "Médio" if eta2H >= 0.06 else "Pequeno"
                            sig_srh    = "✔" if res_srh['p'] < 0.05 else "✘"
                            output += (
                                f"  {fac:<30} {int(res_srh['df']):>4} {res_srh['SS']:>10.3f} "
                                f"{res_srh['H']:>8.3f} {res_srh['p']:>8.4f} {sig_srh:>4} "
                                f"{eta2H:>6.3f} {interp_srh:>10}\n"
                            )
                            if fac == group:
                                srh_group_p      = res_srh['p']
                                srh_group_H      = res_srh['H']
                                srh_group_df     = res_srh['df']
                                srh_group_eta    = eta2H
                                srh_group_interp = interp_srh

                        output += "\n"

                        sig_srh_txt = (
                            "estatisticamente significativo (p < 0,05)"
                            if (not np.isnan(srh_group_p) and srh_group_p < 0.05)
                            else "não significativo (p ≥ 0,05)"
                        )
                        output += (
                            f"  ► O Scheirer-Ray-Hare revelou que o efeito de '{group}' "
                            f"sobre '{target}' foi {sig_srh_txt} "
                            f"[H({int(srh_group_df)})={srh_group_H:.3f}, p={srh_group_p:.4f}]. "
                            f"O η²H de {srh_group_eta:.3f} indica efeito {srh_group_interp.lower()}. "
                        )
                        if moment_col and not np.isnan(srh_inter_p):
                            inter_txt = (
                                f"significativa (p={srh_inter_p:.4f}), sugerindo que o efeito de "
                                f"'{group}' sobre '{target}' varia conforme '{moment_col}'"
                                if srh_inter_p < 0.05 else
                                f"não significativa (p={srh_inter_p:.4f}), indicando efeitos independentes"
                            )
                            output += f"A interação '{group} × {moment_col}' foi {inter_txt}. "
                        output += "\n\n"

                        # Post-hoc Dunn + Bonferroni
                        if not np.isnan(srh_group_p) and srh_group_p < 0.05 and n_groups > 2:
                            output += "━"*75 + "\n"
                            output += "4. ANÁLISE POST-HOC – Dunn + Bonferroni (Scheirer-Ray-Hare)\n"
                            output += "━"*75 + "\n\n"
                            try:
                                from scikit_posthocs import posthoc_dunn
                                dunn = posthoc_dunn(
                                    data, val_col=target, group_col=group, p_adjust='bonferroni'
                                )
                                output += "▸ Dunn + Bonferroni:\n"
                                output += dunn.to_string() + "\n\n"
                                output += (
                                    f"  ► Valores p < 0,05 indicam diferença significativa entre "
                                    f"o par em '{target}'.\n\n"
                                )
                            except ImportError:
                                output += "▸ Mann-Whitney pairwise + Bonferroni (fallback):\n"
                                n_comp_srh = n_groups * (n_groups - 1) // 2
                                for i_g, g1 in enumerate(groups_list):
                                    for g2 in groups_list[i_g+1:]:
                                        s1_ph = data[data[group] == g1][target].values
                                        s2_ph = data[data[group] == g2][target].values
                                        u_ph, p_raw_ph = stats.mannwhitneyu(
                                            s1_ph, s2_ph, alternative='two-sided'
                                        )
                                        p_adj_ph = min(p_raw_ph * n_comp_srh, 1.0)
                                        r_rb_ph  = 1 - (2 * u_ph / (len(s1_ph) * len(s2_ph)))
                                        interp_r_ph = "Grande" if abs(r_rb_ph) >= 0.5 else "Médio" if abs(r_rb_ph) >= 0.3 else "Pequeno"
                                        sig_ph = "✔" if p_adj_ph < 0.05 else "✘"
                                        output += (
                                            f"  {sig_ph} '{g1}' vs '{g2}': "
                                            f"U={u_ph:.1f}, p_ajust={p_adj_ph:.4f}, "
                                            f"r={r_rb_ph:.3f} ({interp_r_ph})\n"
                                        )
                                output += "\n"

            # ════════════════════════════════════════════════════════════════════
            # RAMO C – Comparação simples (sem ID, sem momento, sem covariáveis)
            # Normal 2 gr → t Student/Welch  | Normal ≥3 gr → ANOVA 1-way + Tukey
            # Não-norm 2 gr → Mann-Whitney   | Não-norm ≥3 gr → Kruskal-Wallis + Dunn
            # ════════════════════════════════════════════════════════════════════
            else:
                if is_normal:
                    if n_groups == 2:
                        # ── t de Student ou Welch ─────────────────────────────────
                        res       = stats.ttest_ind(samples[0], samples[1], equal_var=homogeneous)
                        test_name = "t de Student" if homogeneous else "t de Welch (variâncias desiguais)"
                        p_val     = res.pvalue

                        n1, n2 = len(samples[0]), len(samples[1])
                        s1, s2 = np.std(samples[0], ddof=1), np.std(samples[1], ddof=1)

                        # d de Cohen: pooled para variâncias iguais, média quadrática para Welch
                        if homogeneous:
                            sp   = np.sqrt(((n1-1)*s1**2 + (n2-1)*s2**2) / (n1+n2-2))
                            sp_d = sp
                            se_diff = sp * np.sqrt(1/n1 + 1/n2)
                            df_t    = float(n1 + n2 - 2)
                        else:
                            sp_d    = np.sqrt((s1**2 + s2**2) / 2)
                            se_diff = np.sqrt(s1**2/n1 + s2**2/n2)
                            df_t    = (s1**2/n1 + s2**2/n2)**2 / \
                                      ((s1**2/n1)**2/(n1-1) + (s2**2/n2)**2/(n2-1))

                        d       = (np.mean(samples[0]) - np.mean(samples[1])) / sp_d
                        interp_d = "Grande" if abs(d) >= 0.8 else "Médio" if abs(d) >= 0.5 else "Pequeno"

                        t_crit = stats.t.ppf(0.975, df_t)
                        diff   = np.mean(samples[0]) - np.mean(samples[1])
                        ic_lo  = diff - t_crit * se_diff
                        ic_hi  = diff + t_crit * se_diff

                        sig_t = (
                            "estatisticamente significativa (p < 0,05)"
                            if p_val < 0.05 else "não significativa (p ≥ 0,05)"
                        )
                        output += (
                            f"▸ Modelo Selecionado: {test_name}\n"
                            f"  Justificativa: Dois grupos independentes com distribuição normal "
                            f"{'e variâncias homogêneas.' if homogeneous else '(variâncias heterogêneas → correção de Welch aplicada).'}\n\n"
                            f"  t({df_t:.2f}) = {res.statistic:.3f}, p = {res.pvalue:.4f}\n"
                            f"  Diferença de médias: {diff:.3f} (IC 95%: [{ic_lo:.3f}; {ic_hi:.3f}])\n"
                            f"  d de Cohen = {abs(d):.3f} ({interp_d})\n\n"
                            f"  ► A diferença em '{target}' entre '{groups_list[0]}' "
                            f"(M={np.mean(samples[0]):.2f}, DP={np.std(samples[0], ddof=1):.2f}) "
                            f"e '{groups_list[1]}' (M={np.mean(samples[1]):.2f}, DP={np.std(samples[1], ddof=1):.2f}) "
                            f"foi {sig_t} [{test_name}: t({df_t:.2f})={res.statistic:.3f}, p={res.pvalue:.4f}]. "
                            f"O d de Cohen de {abs(d):.3f} indica efeito {interp_d.lower()} "
                            f"(IC 95%: [{ic_lo:.3f}; {ic_hi:.3f}]).\n\n"
                        )

                    else:
                        # ── ANOVA 1-Way ───────────────────────────────────────────
                        res    = stats.f_oneway(*samples)
                        p_val  = res.pvalue
                        n_tot  = len(data[target])
                        grand  = np.mean(data[target])
                        ss_bet = sum(len(s) * (np.mean(s) - grand)**2 for s in samples)
                        ss_tot = sum((x - grand)**2 for x in data[target])
                        ss_wit = ss_tot - ss_bet
                        df_bet = n_groups - 1
                        df_wit = n_tot - n_groups
                        eta2   = ss_bet / ss_tot
                        omega2 = max(
                            (ss_bet - df_bet * (ss_wit / df_wit)) / (ss_tot + ss_wit / df_wit), 0
                        )
                        interp_eta = "Grande" if eta2 >= 0.14 else "Médio" if eta2 >= 0.06 else "Pequeno"
                        sig_f = (
                            "estatisticamente significativa (p < 0,05)"
                            if p_val < 0.05 else "não significativa (p ≥ 0,05)"
                        )
                        output += (
                            f"▸ Modelo Selecionado: ANOVA 1-Way\n"
                            f"  Justificativa: Mais de dois grupos independentes com distribuição normal.\n\n"
                            f"  F({df_bet},{df_wit}) = {res.statistic:.3f}, p = {res.pvalue:.4f}\n"
                            f"  η² = {eta2:.3f} ({interp_eta}), ω² = {omega2:.3f}\n\n"
                            f"  ► A ANOVA 1-Way indicou diferença {sig_f} em '{target}' "
                            f"entre os grupos de '{group}' "
                            f"[F({df_bet},{df_wit})={res.statistic:.3f}, p={res.pvalue:.4f}]. "
                            f"O η² de {eta2:.3f} indica efeito {interp_eta.lower()} "
                            f"({eta2*100:.1f}% da variância explicada).\n\n"
                        )
                        if p_val < 0.05:
                            posthoc_needed = True

                else:  # não-paramétrico
                    if n_groups == 2:
                        # ── Mann-Whitney U ────────────────────────────────────────
                        res    = stats.mannwhitneyu(samples[0], samples[1], alternative='two-sided')
                        p_val  = res.pvalue
                        r_rb   = 1 - (2 * res.statistic / (len(samples[0]) * len(samples[1])))
                        interp_r = "Grande" if abs(r_rb) >= 0.5 else "Médio" if abs(r_rb) >= 0.3 else "Pequeno"
                        sig_u  = (
                            "estatisticamente significativa (p < 0,05)"
                            if p_val < 0.05 else "não significativa (p ≥ 0,05)"
                        )
                        output += (
                            f"▸ Modelo Selecionado: Mann-Whitney U (não-paramétrico)\n"
                            f"  Justificativa: Dois grupos independentes com distribuição não-normal.\n\n"
                            f"  U = {res.statistic:.1f}, p = {res.pvalue:.4f}\n"
                            f"  r Bisserial de Posto = {r_rb:.3f} ({interp_r})\n\n"
                            f"  ► A diferença em '{target}' entre "
                            f"'{groups_list[0]}' (Md={np.median(samples[0]):.2f}) e "
                            f"'{groups_list[1]}' (Md={np.median(samples[1]):.2f}) "
                            f"foi {sig_u} [U={res.statistic:.1f}, p={res.pvalue:.4f}]. "
                            f"O r bisserial de {r_rb:.3f} indica efeito {interp_r.lower()}.\n\n"
                        )

                    else:
                        # ── Kruskal-Wallis ────────────────────────────────────────
                        res    = stats.kruskal(*samples)
                        p_val  = res.pvalue
                        n_tot  = len(data)
                        eps2   = max(res.statistic / (n_tot - 1), 0)
                        interp_eps = "Grande" if eps2 >= 0.14 else "Médio" if eps2 >= 0.06 else "Pequeno"
                        sig_k  = (
                            "estatisticamente significativa (p < 0,05)"
                            if p_val < 0.05 else "não significativa (p ≥ 0,05)"
                        )
                        output += (
                            f"▸ Modelo Selecionado: Kruskal-Wallis (não-paramétrico)\n"
                            f"  Justificativa: Mais de dois grupos independentes com distribuição não-normal.\n\n"
                            f"  H({n_groups-1}) = {res.statistic:.3f}, p = {res.pvalue:.4f}\n"
                            f"  ε² = {eps2:.3f} ({interp_eps})\n\n"
                            f"  ► O Kruskal-Wallis indicou diferença {sig_k} em '{target}' "
                            f"entre os grupos de '{group}' "
                            f"[H({n_groups-1})={res.statistic:.3f}, p={res.pvalue:.4f}]. "
                            f"O ε² de {eps2:.3f} indica efeito {interp_eps.lower()}.\n\n"
                        )
                        if p_val < 0.05:
                            posthoc_needed = True

            # ── 7. POST-HOC – Ramo C (ANOVA 1-way ou Kruskal-Wallis significativos) ──
            if posthoc_needed:
                output += "━"*75 + "\n"
                output += "4. ANÁLISE POST-HOC (Comparações Múltiplas)\n"
                output += "━"*75 + "\n\n"

                if is_normal:
                    from statsmodels.stats.multicomp import pairwise_tukeyhsd
                    tukey = pairwise_tukeyhsd(data[target], data[group], alpha=0.05)
                    output += "▸ Tukey HSD (pós ANOVA 1-Way):\n"
                    output += str(tukey.summary()) + "\n\n"
                    output += "  d de Cohen pairwise:\n"
                    for i, g1 in enumerate(groups_list):
                        for g2 in groups_list[i+1:]:
                            s1 = data[data[group] == g1][target].values
                            s2 = data[data[group] == g2][target].values
                            sp = np.sqrt(
                                ((len(s1)-1)*np.std(s1, ddof=1)**2 +
                                 (len(s2)-1)*np.std(s2, ddof=1)**2) / (len(s1)+len(s2)-2)
                            )
                            d  = abs(np.mean(s1) - np.mean(s2)) / sp if sp > 0 else 0
                            interp_d = "Grande" if d >= 0.8 else "Médio" if d >= 0.5 else "Pequeno"
                            output += f"  • '{g1}' vs '{g2}': d={d:.3f} ({interp_d})\n"
                    output += (
                        f"\n  ► Pares com p ajustado < 0,05 diferem "
                        f"significativamente em '{target}'.\n\n"
                    )

                else:
                    try:
                        from scikit_posthocs import posthoc_dunn
                        dunn = posthoc_dunn(data, val_col=target, group_col=group, p_adjust='bonferroni')
                        output += "▸ Dunn + Bonferroni (pós Kruskal-Wallis):\n"
                        output += dunn.to_string() + "\n\n"
                        output += (
                            f"  ► Valores p < 0,05 indicam diferença significativa "
                            f"entre o par em '{target}'.\n\n"
                        )
                    except ImportError:
                        output += "▸ Mann-Whitney pairwise + Bonferroni (fallback):\n"
                        n_comp = n_groups * (n_groups - 1) // 2
                        for i, g1 in enumerate(groups_list):
                            for g2 in groups_list[i+1:]:
                                s1 = data[data[group] == g1][target].values
                                s2 = data[data[group] == g2][target].values
                                u, p_raw = stats.mannwhitneyu(s1, s2, alternative='two-sided')
                                p_adj    = min(p_raw * n_comp, 1.0)
                                r_rb     = 1 - (2 * u / (len(s1) * len(s2)))
                                interp_r = "Grande" if abs(r_rb) >= 0.5 else "Médio" if abs(r_rb) >= 0.3 else "Pequeno"
                                sig_pair = "✔" if p_adj < 0.05 else "✘"
                                output += (
                                    f"  {sig_pair} '{g1}' vs '{g2}': "
                                    f"U={u:.1f}, p_ajust={p_adj:.4f}, r={r_rb:.3f} ({interp_r})\n"
                                )
                        output += "\n"

            # ── 8. RODAPÉ ─────────────────────────────────────────────────────────
            section_num = "5" if posthoc_needed else "4"
            output += "━"*75 + "\n"
            output += f"{section_num}. CITAÇÃO DO SOFTWARE e NOTA METODOLÓGICA\n"
            output += "━"*75 + "\n\n"
            output += (
                "CITAÇÃO DO SOFTWARE: NEVES, Eduardo Borba. Software de análise estatística para comparação de medidas entre grupos. Versão 1.01. [S.l.]: Zenodo, 2026. Software. DOI:https://doi.org/10.5281/zenodo.19713014 \n\n\n"
                "Os testes foram selecionados com base na verificação dos pressupostos de normalidade (Shapiro-Wilk) e homogeneidade de variâncias (Levene).\n"
                "Tamanhos de efeito foram calculados conforme Cohen (1988) e Field (2018).\n"
                "COHEN, J. Statistical power analysis for the behavioral sciences. 2. ed. Hillsdale, NJ: Lawrence Erlbaum Associates, 1988.\n" 
                "FIELD, A. Discovering statistics using IBM SPSS statistics. 5. ed. London: SAGE Publications, 2018.\n"
                "Nível de significância adotado: α = 0,05.\n\n"
            )

            self.text_result.insert("end", output)

            # ── 9. GRÁFICOS ───────────────────────────────────────────────────────
            sns.set_theme(style="whitegrid", font_scale=1.05)
            palette = "viridis"

            # Gráfico 1: Boxplot + stripplot
            fig1, ax1 = plt.subplots(figsize=(8, 5))
            hue_box = moment_col if moment_col else group
            sns.boxplot(
                x=group, y=target, hue=hue_box, data=data, ax=ax1,
                palette=palette, width=0.5, fliersize=0, legend=bool(moment_col)
            )
            sns.stripplot(
                x=group, y=target, hue=hue_box, data=data, ax=ax1,
                palette=palette, dodge=bool(moment_col), alpha=0.45,
                size=5, jitter=True, legend=False
            )
            ax1.set_title(f"Distribuição de '{target}' por '{group}'", fontweight='bold', pad=12)
            ax1.set_xlabel(group)
            ax1.set_ylabel(target)
            if moment_col and ax1.get_legend():
                ax1.legend(title=moment_col, bbox_to_anchor=(1.01, 1), loc='upper left')
            fig1.tight_layout()
            self.figs.append(fig1)

            # Gráfico 2: Perfil de médias com IC 95%
            fig2, ax2 = plt.subplots(figsize=(8, 5))
            if moment_col:
                sns.pointplot(
                    x=moment_col, y=target, hue=group, data=data, ax=ax2,
                    capsize=.12, err_kws={'linewidth': 1.8}, palette=palette,
                    markers=['o','s','D','P'][:n_groups],
                    linestyles=['-','--',':','-.'][:n_groups]
                )
                ax2.set_title(f"Evolução de '{target}' por '{group}' e '{moment_col}'",
                              fontweight='bold', pad=12)
                ax2.set_xlabel(moment_col)
                ax2.legend(title=group, bbox_to_anchor=(1.01, 1), loc='upper left')
            else:
                ci_data = []
                for g in groups_list:
                    s    = data[data[group] == g][target]
                    m    = s.mean()
                    se   = s.sem()
                    t_c  = stats.t.ppf(0.975, len(s) - 1)
                    ci_data.append({'Grupo': g, 'Média': m,
                                    'IC_lo': m - t_c*se, 'IC_hi': m + t_c*se})
                df_ci  = pd.DataFrame(ci_data)
                colors = sns.color_palette(palette, n_groups)
                for i, row in df_ci.iterrows():
                    ax2.bar(row['Grupo'], row['Média'], color=colors[i], alpha=0.75, width=0.5)
                    ax2.errorbar(
                        row['Grupo'], row['Média'],
                        yerr=[[row['Média']-row['IC_lo']], [row['IC_hi']-row['Média']]],
                        fmt='none', color='black', capsize=6, linewidth=2
                    )
                    ax2.text(i, row['IC_hi'] + 0.01*(data[target].max()-data[target].min()),
                             f"M={row['Média']:.2f}", ha='center', va='bottom', fontsize=9)
                ax2.set_title(f"Médias ± IC 95% de '{target}' por '{group}'",
                              fontweight='bold', pad=12)
                ax2.set_xlabel(group)
            ax2.set_ylabel(target)
            fig2.tight_layout()
            self.figs.append(fig2)

            # Gráfico 3: Violin plot
            fig3, ax3 = plt.subplots(figsize=(8, 5))
            sns.violinplot(
                x=group, y=target,
                hue=moment_col if moment_col else group,
                data=data, ax=ax3, palette=palette,
                inner='quartile', split=False, legend=bool(moment_col)
            )
            ax3.set_title(f"Forma da Distribuição de '{target}' por '{group}'",
                          fontweight='bold', pad=12)
            ax3.set_xlabel(group)
            ax3.set_ylabel(target)
            if moment_col and ax3.get_legend():
                ax3.legend(title=moment_col, bbox_to_anchor=(1.01, 1), loc='upper left')
            fig3.tight_layout()
            self.figs.append(fig3)

            # Gráfico 4: Q-Q plots por grupo
            n_cols_qq = min(n_groups, 3)
            n_rows_qq = (n_groups + n_cols_qq - 1) // n_cols_qq
            fig4, axes4 = plt.subplots(n_rows_qq, n_cols_qq,
                                        figsize=(5*n_cols_qq, 4*n_rows_qq), squeeze=False)
            for idx, g in enumerate(groups_list):
                ax_qq = axes4[idx // n_cols_qq][idx % n_cols_qq]
                s     = data[data[group] == g][target]
                (osm, osr), (slope, intercept, _) = stats.probplot(s, dist="norm")
                ax_qq.plot(osm, osr, 'o', alpha=0.6, markersize=4,
                           color=sns.color_palette(palette, n_groups)[idx])
                ax_qq.plot(osm, slope*np.array(osm)+intercept, 'r-', linewidth=1.5)
                p_sh_g = sw_results.get(g, (None, None, None))[1]
                p_txt  = f"SW p={p_sh_g:.4f}" if p_sh_g is not None else ""
                ax_qq.set_title(f"Q-Q: '{g}'\n{p_txt}", fontsize=10)
                ax_qq.set_xlabel("Quantis Teóricos")
                ax_qq.set_ylabel("Quantis Observados")
            for idx in range(n_groups, n_rows_qq * n_cols_qq):
                axes4[idx // n_cols_qq][idx % n_cols_qq].set_visible(False)
            fig4.suptitle(f"Gráficos Q-Q Normal — '{target}'", fontweight='bold', y=1.01)
            fig4.tight_layout()
            self.figs.append(fig4)

            # Renderizar gráficos
            for f in self.figs:
                canvas = tkagg.FigureCanvasTkAgg(f, master=self.result_frame)
                canvas.draw()
                canvas.get_tk_widget().pack(fill="x", pady=20, padx=10)

        except Exception as e:
            import traceback
            messagebox.showerror("Erro de Processamento",
                                 f"Ocorreu um erro:\n{str(e)}\n\n{traceback.format_exc()}")


    def export_results(self):
        # Verifica se existem figuras ou texto para exportar
        if not self.figs and len(self.text_result.get("1.0", "end").strip()) == 0:
            messagebox.showwarning("Aviso", "Não existem resultados para exportar.")
            return
            
        # Abre a caixa de diálogo para salvar o ficheiro
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Arquivo de Texto", "*.txt"), ("Todos os arquivos", "*.*")]
        )
        
        if path:
            try:
                # 1. Guarda o relatório de texto
                with open(path, "w", encoding="utf-8") as f:
                    f.write(self.text_result.get("1.0", "end"))
                
                # 2. Guarda as figuras individualmente
                # Usa o caminho do ficheiro de texto para criar os nomes das imagens
                base_path = os.path.splitext(path)[0]
                
                for i, fig in enumerate(self.figs):
                    # Define um nome como "nome_do_arquivo_grafico_1.png"
                    fig_name = f"{base_path}_grafico_{i+1}.png"
                    # Guarda a figura com alta resolução (300 DPI) e remove bordas excessivas
                    fig.savefig(fig_name, dpi=300, bbox_inches='tight')
                
                messagebox.showinfo("Sucesso", 
                    f"Exportação concluída!\n\n- Relatório: {os.path.basename(path)}\n"
                    f"- Gráficos: {len(self.figs)} ficheiros .png gerados.")
                    
            except Exception as e:
                messagebox.showerror("Erro na Exportação", f"Ocorreu uma falha ao salvar os ficheiros: {e}")

    def safe_exit(self):
        plt.close('all'); self.quit(); self.destroy()

if __name__ == "__main__":
    StatisticalApp().mainloop()


# In[ ]:





# In[ ]:





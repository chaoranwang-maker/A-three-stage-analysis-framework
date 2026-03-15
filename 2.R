
packages <- c("grf", "tidyverse", "ggplot2", "readxl", "stringr", "mice", 
              "cobalt", "sensitivitymv", "survival", "survminer", "WeightIt", 
              "tableone", "Hmisc", "survey", "gridExtra")
installed <- rownames(installed.packages())
for (pkg in packages) {
  if (!pkg %in% installed) install.packages(pkg)
}
library(grf)
library(tidyverse)
library(ggplot2)
library(readxl)
library(stringr)
library(mice)
library(cobalt)
library(sensitivitymv)
library(survival)
library(survminer)
library(WeightIt)
library(tableone)
library(Hmisc)
library(survey)
library(gridExtra)

save_path <- ""
if (!dir.exists(save_path)) dir.create(save_path, recursive = TRUE)

df_raw <- read_excel(
  path = "",
  na = c("", "NA", "N/A", ".", "NULL")
)

cat("Original data dimensions:", nrow(df_raw), "rows x", ncol(df_raw), "columns\n")



df <- df_raw %>% mutate(
  
  age = as.numeric(`年龄（非入院）`),
  
  
  sex = as.numeric(`性别（男1女2）`),
  
  
  age_score = case_when(
    age < 40 ~ 0,
    age >= 40 & age < 50 ~ 1,
    age >= 50 & age < 60 ~ 2,
    age >= 60 & age < 70 ~ 3,
    age >= 70 & age < 80 ~ 4,
    age >= 80 ~ 5,
    TRUE ~ NA_real_
  ),
  original_cci = as.numeric(`CCI合并症`),
  CCI_score = original_cci + age_score,
  
  
  smoke_drink_raw = as.numeric(`吸烟饮酒史（无：0，吸烟；1，饮酒：2，均有：3`),
  smoke = ifelse(smoke_drink_raw %in% c(1, 3), 1, 0),
  drink = ifelse(smoke_drink_raw %in% c(2, 3), 1, 0),
  
  
  vaccine_num = as.numeric(`疫苗接种次数`),
  

  discharge_type = as.numeric(`出院分型`),
  

  treatment = as.numeric(`是否用药`),
  
  
  anti_viral = as.numeric(`是否联合抗病毒`),
  
  
  admission_temp = as.numeric(`入院体温`),
  
 
  treatment_duration = as.numeric(`用药时间`)
)


df$vaccine_num[is.na(df$vaccine_num)] <- 10


df$vaccine <- factor(df$vaccine_num, 
                     levels = c(0,1,2,3,10),
                     labels = c("Unvaccinated", "Dose1", "Dose2", "Dose3plus", "Unknown"))
cat("Variable conversion check:\n")
cat("Age range:", range(df$age, na.rm = TRUE), "\n")
cat("Sex distribution:"); print(table(df$sex, useNA = "ifany"))
cat("CCI_score range:", range(df$CCI_score, na.rm = TRUE), "\n")
cat("Smoke/Drink cross-tab:\n"); print(table(smoke=df$smoke, drink=df$drink, useNA = "ifany"))
cat("Vaccine distribution:"); print(table(df$vaccine, useNA = "ifany"))
cat("Discharge_type:"); print(table(df$discharge_type, useNA = "ifany"))
cat("Treatment:"); print(table(df$treatment, useNA = "ifany"))
cat("Antiviral:"); print(table(df$anti_viral, useNA = "ifany"))
df <- df %>% rename(`Clinical_classification` = discharge_type)


lab_vars_original <- c("WBC", "NEUT%", "NE#", "LYM%", "LY#", "MONO%", "MO#", "EOS%", "EO#", "BASO%", "BA#",
                       "RBC", "HGB", "HCT", "MCV", "MCH", "RDW-SD", "RDW-CV", "MCHC", "PLT", "MPV",
                       "Plateletcrit", "PLCR", "PDW", "LDH", "CK", "CK-MB", "HBDH", "CRP", "IL-6", "SAA", "PCT",
                       "ALT", "AST", "TBIL", "DBIL", "TP", "ALB", "GLO", "A/G", "CHE",
                       "K", "Na", "Cl", "CA", "MG", "PHOS", "UREA", "CREA", "URCA", "GLU",
                       "TCO2", "AG", "Mosm", "eGFR",
                       "PT", "PTA", "APTT", "Fb", "INR", "FDP", "DD", "TT",
                       "CD3+/CD45+", "CD3+", "CD3+CD8+/CD45+", "CD3+CD8+",
                       "CD3+CD4+/CD45+", "CD3+CD4+", "CD45+", "Ratio",
                       "CD16+CD56+/CD45+", "CD16+CD56+", "CD19+/CD45+", "CD19+")

lab_vars_clean <- make.names(lab_vars_original, unique = TRUE)
lab_mapping <- setNames(lab_vars_clean, lab_vars_original)

for (i in seq_along(lab_vars_original)) {
  orig_name <- lab_vars_original[i]
  clean_name <- lab_vars_clean[i]
  if (orig_name %in% colnames(df_raw)) {
    df[[clean_name]] <- as.numeric(df_raw[[orig_name]])
  }
}

available_lab_vars <- intersect(lab_vars_clean, colnames(df))
cat("Available laboratory indicators:", length(available_lab_vars), "\n")

OBSERVATION_PERIOD <- 28

df <- df %>% mutate(
  remission_time_raw = as.numeric(`全部症状缓解时间`),
  remission_time = ifelse(remission_time_raw > OBSERVATION_PERIOD | remission_time_raw == 1000,
                          OBSERVATION_PERIOD, remission_time_raw),
  event = ifelse(remission_time_raw > OBSERVATION_PERIOD | remission_time_raw == 1000, 0, 1),
  time = remission_time
)

censored_count <- sum(df$remission_time_raw > OBSERVATION_PERIOD | df$remission_time_raw == 1000, na.rm = TRUE)
cat("28-day truncation:\n")
cat("  Non-remission/censored:", censored_count, "\n")
cat("  Remission time range:", range(df$remission_time, na.rm = TRUE), "days\n")
cat("  Events (remission):", sum(df$event), "\n")
cat("\n=== Step 5: Missing Data Handling ===\n")

basic_vars <- c("age", "sex", "CCI_score", "smoke", "drink", "vaccine", 
                "Clinical_classification", "anti_viral", "admission_temp",
                "remission_time", "event")

cat("Basic variable missing check:\n")
for (var in basic_vars) {
  n_missing <- sum(is.na(df[[var]]))
  cat(sprintf("  %s: %d missing (%.2f%%)\n", var, n_missing, 100*n_missing/nrow(df)))
  if (n_missing > 0) {
    stop(paste("Error: Basic variable", var, "has missing values!"))
  }
}

cat("\nProcessing lab indicators only...\n")

lab_missing_rate <- sapply(available_lab_vars, function(v) mean(is.na(df[[v]])))
cat("Lab indicators >20% missing:\n")
print(sort(lab_missing_rate[lab_missing_rate > 0.2], decreasing = TRUE))

high_missing_labs <- names(lab_missing_rate[lab_missing_rate > 0.4])
if (length(high_missing_labs) > 0) {
  cat("Deleting high-missing labs(>40%):", paste(high_missing_labs, collapse = ", "), "\n")
  available_lab_vars <- setdiff(available_lab_vars, high_missing_labs)
}

labs_to_impute <- available_lab_vars[sapply(available_lab_vars, function(v) any(is.na(df[[v]])))]

if (length(labs_to_impute) > 0) {
  cat("\nLabs requiring imputation:", paste(labs_to_impute, collapse = ", "), "\n")
  
  impute_data <- df %>% 
    select(all_of(labs_to_impute)) %>%
    mutate(across(everything(), as.numeric))
  
  cat("Performing MICE (m=5)...\n")
  imp <- mice(impute_data, m = 5, method = "pmm", maxit = 10, seed = 123, printFlag = FALSE)
  completed_labs <- complete(imp, action = 1)
  
  for (col in names(completed_labs)) {
    df[[col]] <- completed_labs[[col]]
  }
  cat("Imputation completed\n")
} else {
  cat("All lab indicators complete\n")
}


X_screen <- df %>% 
  select(age, sex, CCI_score, smoke, drink, vaccine_num, 
         Clinical_classification, anti_viral, admission_temp,
         all_of(available_lab_vars)) %>%
  mutate(across(everything(), as.numeric)) %>%
  as.matrix()

Y <- df$remission_time
T <- df$treatment

set.seed(123)
cf_screen <- causal_forest(
  X = X_screen,
  Y = Y,
  W = T,
  num.trees = 2000,
  min.node.size = max(5, floor(nrow(X_screen) * 0.01)),
  mtry = max(1, floor(sqrt(ncol(X_screen)))),
  sample.fraction = 0.5,
  honesty = TRUE
)


importance_all <- variable_importance(cf_screen)


n_confounders <- 9  # age, sex, CCI_score, smoke, drink, vaccine_num, clinical_classification, anti_viral, admission_temp
lab_importance <- importance_all[-(1:n_confounders)]

screen_df <- data.frame(
  Variable = colnames(X_screen)[-(1:n_confounders)],
  Importance = as.vector(lab_importance),
  stringsAsFactors = FALSE
) %>% arrange(desc(Importance))

core_markers <- screen_df %>% head(10) %>% pull(Variable)
cat("Top 10 Core Markers (confounder-adjusted screening):\n")
cat(paste(core_markers, collapse = ", "), "\n")




df <- df %>% mutate(
  sex_f = factor(sex, levels = c(1, 2), labels = c("Male", "Female")),
  smoke_f = factor(smoke, levels = c(0, 1), labels = c("No", "Yes")),
  drink_f = factor(drink, levels = c(0, 1), labels = c("No", "Yes")),
  
 
  vaccine_f = fct_recode(vaccine,
                         "1 dose" = "Dose1",
                         "2 doses" = "Dose2",
                         "3+ doses" = "Dose3plus"
                         
  ),
  
  Clinical_classification_f = factor(Clinical_classification,
                                     levels = c(1, 2, 3, 4),
                                     labels = c("Mild", "Moderate", "Severe", "Critical")),
  anti_viral_f = factor(anti_viral, levels = c(0, 1), labels = c("No", "Yes")),
  treatment_f = factor(treatment, levels = c(0, 1), labels = c("Control", "Yindan"))
)

df_table1 <- df %>%
  mutate(
    treatment_label = factor(treatment, levels = c(0, 1), labels = c("Control", "Yindan"))
  )

vars_table1 <- c(
  "age", "sex_f", "CCI_score", "smoke_f", "drink_f",
  "vaccine_f", "Clinical_classification_f", "anti_viral_f", "admission_temp",
  available_lab_vars
)

vars_table1 <- intersect(vars_table1, colnames(df_table1))
cat_vars <- c("sex_f", "smoke_f", "drink_f", "vaccine_f", "Clinical_classification_f", "anti_viral_f")
nonnormal_vars <- c("age", "CCI_score", "admission_temp", available_lab_vars)

table1 <- CreateTableOne(
  vars = vars_table1,
  strata = "treatment_label",
  data = df_table1,
  factorVars = intersect(cat_vars, vars_table1),
  includeNA = TRUE,
  test = TRUE
)

cat("\nBaseline Table (Unweighted):\n")
print(table1, 
      nonnormal = intersect(nonnormal_vars, vars_table1),
      exact = intersect(cat_vars, vars_table1),
      showAllLevels = TRUE,
      smd = TRUE,
      formatOptions = list(big.mark = ",", digits = 2))

table1_for_save <- print(table1, 
                         nonnormal = intersect(nonnormal_vars, vars_table1),
                         exact = intersect(cat_vars, vars_table1),
                         showAllLevels = TRUE,
                         smd = TRUE,
                         formatOptions = list(big.mark = ",", digits = 2),
                         noSpaces = TRUE,
                         printToggle = FALSE)

write.csv(as.data.frame(table1_for_save), 
          paste0(save_path, "Table1_Baseline_Complete.csv"), 
          row.names = TRUE,
          fileEncoding = "UTF-8")

cat("\nBaseline table saved to:", paste0(save_path, "Table1_Baseline_Complete.csv"), "\n")


weighting_vars <- c("age", "sex", "CCI_score", "smoke", "drink",
                    "vaccine", "Clinical_classification", "anti_viral", "admission_temp",
                    core_markers)

weighting_formula <- as.formula(
  paste("treatment ~", paste(intersect(weighting_vars, colnames(df)), collapse = " + "))
)

cat("IPTW formula:"); print(weighting_formula)

set.seed(456)
W.out <- weightit(weighting_formula,
                  data = df,
                  method = "gbm",
                  estimand = "ATE",
                  stop.method = "es.mean")

cat("\nWeight distribution:\n"); print(summary(W.out$weights))

weight_q <- quantile(W.out$weights, probs = c(0.01, 0.99), na.rm = TRUE)
df$iptw_weights <- W.out$weights
df$iptw_weights_stab <- pmin(pmax(df$iptw_weights, weight_q[1]), weight_q[2])

cat(paste0("\nStabilized weights (1%-99%): ", round(min(df$iptw_weights_stab), 3), " - ", 
           round(max(df$iptw_weights_stab), 3), "\n"))

cat("\nSMD after weighting:\n")
print(bal.tab(W.out, un = TRUE, thresholds = c(m = 0.1)))

p_love <- love.plot(W.out, 
                    threshold = 0.1, 
                    var.order = "unadjusted", 
                    abs = TRUE,
                    stars = "raw",
                    title = "Covariate Balance: Before vs After IPTW Weighting",
                    subtitle = "Starred points indicate raw mean differences")
ggsave(paste0(save_path, "IPTW_LovePlot.png"), p_love, width = 10, height = 10, dpi = 300)



fit_weighted <- survfit(Surv(time, event) ~ treatment_f, 
                        data = df, weights = iptw_weights_stab)

cox_weighted <- coxph(Surv(time, event) ~ treatment_f, 
                      data = df, weights = iptw_weights_stab, robust = TRUE)

cat("\nIPTW-Weighted Cox:\n"); print(summary(cox_weighted))

hr <- exp(coef(cox_weighted))
hr_ci <- exp(confint(cox_weighted))
p_val <- summary(cox_weighted)$coefficients[1, "Pr(>|z|)"]

if (is.matrix(hr_ci)) {
  ci_lower <- hr_ci[1, 1]
  ci_upper <- hr_ci[1, 2]
} else {
  ci_lower <- hr_ci[1]
  ci_upper <- hr_ci[2]
}

cat(paste0("\nHR = ", round(hr, 2), " (95%CI: ", round(ci_lower, 2), "-", round(ci_upper, 2), ")\n"))
med_surv <- surv_median(fit_weighted)
cat("Median remission time - Control:", ifelse(is.na(med_surv$median[1]), ">28", round(med_surv$median[1],1)),
    " Yindan:", ifelse(is.na(med_surv$median[2]), ">28", round(med_surv$median[2],1)), "\n")

write.csv(data.frame(HR=hr, HR_lower=ci_lower, HR_upper=ci_upper, P=p_val),
          paste0(save_path, "IPTW_Cox_Results.csv"))



if(!"vaccine_num" %in% names(df)) {
  df$vaccine_num <- as.numeric(as.character(df$vaccine))
  df$vaccine_num[is.na(df$vaccine_num)] <- 10
}


W_matrix_temp <- df %>% 
  select(age, sex, CCI_score, smoke, drink, vaccine_num, Clinical_classification, 
         anti_viral, admission_temp, all_of(core_markers)) %>%
  mutate(across(everything(), as.numeric),
         across(everything(), ~ifelse(is.na(.), median(., na.rm = TRUE), .))) %>% 
  as.matrix()

ps_scores <- predict(regression_forest(W_matrix_temp, df$treatment, num.trees = 2000))$predictions
df$ps_score <- ps_scores


cat("\n--- 9.1 Logistic Propensity Score ---\n")
ps_logit_model <- glm(weighting_formula, data = df, family = binomial(link = "logit"))
df$ps_logit <- predict(ps_logit_model, type = "response")
df$iptw_logit <- ifelse(df$treatment == 1, 1/df$ps_logit, 1/(1-df$ps_logit))
w_q_logit <- quantile(df$iptw_logit, probs = c(0.01, 0.99), na.rm = TRUE)
df$iptw_logit_trunc <- pmin(pmax(df$iptw_logit, w_q_logit[1]), w_q_logit[2])

cox_logit <- coxph(Surv(time, event) ~ treatment_f, 
                   data = df, weights = iptw_logit_trunc, robust = TRUE)
cat("\nLogistic PS IPTW Results:\n")
print(summary(cox_logit))


cat("\n--- 9.2 Overlap Weighting (ATO) ---\n")
W_ato <- weightit(weighting_formula,
                  data = df,
                  method = "gbm",
                  estimand = "ATO",
                  stop.method = "es.mean")
df$weights_ato <- W_ato$weights

cox_ato <- coxph(Surv(time, event) ~ treatment_f, 
                 data = df, weights = weights_ato, robust = TRUE)
cat("\nOverlap Weighting Results:\n")
print(summary(cox_ato))


cat("\n--- 9.3 Trimmed Analysis (Positivity) ---\n")
common_support <- (df$ps_score > 0.05) & (df$ps_score < 0.95)
df_trimmed <- df[common_support, ]

cox_trimmed <- coxph(Surv(time, event) ~ treatment_f, 
                     data = df_trimmed, 
                     weights = iptw_weights_stab[common_support], 
                     robust = TRUE)
cat("\nTrimmed Sample Results:\n")
print(summary(cox_trimmed))

cat("Original sample:", nrow(df), "| After trimming:", nrow(df_trimmed), 
    "(", round(mean(common_support)*100, 1), "%)\n")


sens_summary <- data.frame(
  Method = c("Main (GBM-IPTW)", "Logistic PS", "Overlap Weighting (ATO)", "Trimmed (Positivity)"),
  HR = c(hr, exp(coef(cox_logit)), exp(coef(cox_ato)), exp(coef(cox_trimmed))),
  CI_Lower = c(ci_lower, exp(confint(cox_logit))[1], exp(confint(cox_ato))[1], exp(confint(cox_trimmed))[1]),
  CI_Upper = c(ci_upper, exp(confint(cox_logit))[2], exp(confint(cox_ato))[2], exp(confint(cox_trimmed))[2])
)
write.csv(sens_summary, paste0(save_path, "Table_Sensitivity_Summary.csv"), row.names = FALSE)

cat("\nSensitivity analyses completed\n")


nejm_colors <- c("Control" = "#BC3C29", "YDJDG" = "#0072B5")

fit_summary <- summary(fit_weighted, times = unique(fit_weighted$time))
km_data <- data.frame(
  time = fit_summary$time,
  surv = fit_summary$surv,
  lower = fit_summary$lower,
  upper = fit_summary$upper,
  strata = gsub("Yindan", "YDJDG", gsub("treatment_f=", "", as.character(fit_summary$strata))),
  stringsAsFactors = FALSE
)

censor_summary <- summary(fit_weighted, times = df$time[df$event == 0], extend = TRUE)
censor_data <- data.frame(
  time = censor_summary$time,
  surv = censor_summary$surv,
  strata = gsub("Yindan", "YDJDG", gsub("treatment_f=", "", as.character(censor_summary$strata)))
) %>% distinct()

hr_text <- paste0("HR = ", round(hr, 2), " (95% CI, ", round(ci_lower, 2), "-", round(ci_upper, 2), ")")
p_text <- ifelse(p_val < 0.001, "P < 0.001", paste0("P = ", format(round(p_val, 3), nsmall = 3)))

p_km_main <- ggplot(km_data, aes(x = time, y = surv, color = strata, fill = strata)) +
  geom_ribbon(aes(ymin = lower, ymax = upper), alpha = 0.15, color = NA) +
  geom_step(linewidth = 1.2) +
  geom_point(data = censor_data, aes(x = time, y = surv), shape = 3, size = 3, stroke = 1.2, color = "black") +
  geom_hline(yintercept = 0.5, linetype = "dotted", color = "gray40", linewidth = 0.8) +
  scale_x_continuous(breaks = seq(0, 28, 7), limits = c(0, 28), expand = c(0.01, 0)) +
  scale_y_continuous(labels = scales::percent, limits = c(0, 1), expand = c(0.01, 0)) +
  scale_color_manual(values = nejm_colors) +
  scale_fill_manual(values = nejm_colors) +
  annotate("text", x = 26, y = 0.28, label = paste(hr_text, p_text, sep = "\n"),
           hjust = 1, vjust = 0, size = 5, fontface = "bold", color = "black") +
  labs(x = "Time (days)", y = "Probability of Symptom Relief",
       title = "Time to Symptom Relief (IPTW-Weighted)") +
  theme_classic(base_size = 14) +
  theme(legend.position.inside = c(0.02, 0.02), 
        legend.justification = c("left", "bottom"),
        legend.background = element_blank(), 
        legend.title = element_blank(),
        legend.text = element_text(size = 12, face = "bold"))
print(p_km_main)
ggsave(paste0(save_path, "KM_Main_NejmStyle.png"), p_km_main, width = 10, height = 8, dpi = 600)

time_points <- c(0, 7, 14, 21, 28)
risk_summary <- summary(fit_weighted, times = time_points, extend = TRUE)
risk_df <- data.frame(
  time = risk_summary$time,
  n.risk = risk_summary$n.risk,
  strata = gsub("Yindan", "YDJDG", gsub("treatment_f=", "", as.character(risk_summary$strata)))
) %>% distinct()

risk_wide <- risk_df %>%
  pivot_wider(names_from = strata, values_from = n.risk, values_fill = 0) %>%
  mutate(Time = paste0("Day ", time)) %>%
  select(Time, Control, YDJDG)

cat("\nNumbers at Risk:\n"); print(risk_wide, row.names = FALSE)
write.csv(risk_wide, paste0(save_path, "Numbers_at_Risk.csv"), row.names = FALSE)

p_risk <- ggplot(risk_df, aes(x = time, y = factor(strata, levels = c("YDJDG", "Control")))) +
  geom_text(aes(label = n.risk), size = 4.5, fontface = "bold",
            color = ifelse(risk_df$strata == "Control", nejm_colors["Control"], nejm_colors["YDJDG"])) +
  scale_x_continuous(breaks = time_points, limits = c(0, 28), expand = c(0.01, 0)) +
  labs(x = "Time (days)", y = "", title = "Numbers at Risk (IPTW-weighted)") +
  theme_minimal() + theme(panel.grid = element_blank(), plot.title = element_text(size = 12, face = "bold"))
ggsave(paste0(save_path, "KM_RiskTable.png"), p_risk, width = 10, height = 3, dpi = 600)

W_matrix <- df %>% 
  select(age, sex,vaccine,Clinical_classification,CCI_score, smoke, drink, anti_viral, admission_temp, all_of(core_markers)) %>%
  mutate(across(everything(), ~ ifelse(is.na(.), median(., na.rm = TRUE), .))) %>%
  as.matrix()

ps_forest <- regression_forest(W_matrix, df$treatment, num.trees = 2000)
ps_scores <- predict(ps_forest)$predictions

ps_df <- data.frame(PS = ps_scores, 
                    Treatment = factor(df$treatment, labels = c("Control", "Treated")))

p_ps <- ggplot(ps_df, aes(x = PS, fill = Treatment)) +
  geom_density(alpha = 0.6) +
  geom_vline(xintercept = c(0.05, 0.95), linetype = "dashed", color = "red") +
  scale_fill_manual(values = c("#E69F00", "#56B4E9")) +
  labs(title = "Propensity Score Distribution", x = "PS", y = "Density") +
  theme_classic()
ggsave(paste0(save_path, "PS_Distribution.png"), p_ps, width = 8, height = 6, dpi = 300)

valid_ps <- (ps_scores > 0.05) & (ps_scores < 0.95)
prop_valid <- mean(valid_ps)
cat("Proportion in common support (0.05-0.95):", round(prop_valid*100, 1), "%\n")
cat("  PS < 0.05:", mean(ps_scores < 0.05)*100, "%\n")
cat("  PS > 0.95:", mean(ps_scores > 0.95)*100, "%\n")


W_confounders <- df %>% 
  select(age,sex,vaccine,Clinical_classification, CCI_score, smoke, drink, anti_viral, admission_temp) %>%
  mutate(across(everything(), as.numeric)) %>%
  as.matrix()

X_modifiers <- df %>% 
  select(all_of(available_lab_vars)) %>%
  as.matrix()

Y <- df$remission_time
T <- df$treatment

Y_forest <- regression_forest(W_confounders, Y, num.trees = 2000, sample.weights = df$iptw_weights_stab)
Y_hat <- predict(Y_forest)$predictions
Y_resid <- Y - Y_hat

cat("Residualization completed\n")


set.seed(123)
n <- nrow(X_modifiers)
p <- ncol(X_modifiers)

cf_model <- causal_forest(
  X = X_modifiers,
  Y = Y_resid,
  W = T,
  num.trees = 4000,
  min.node.size = max(5, floor(n * 0.01)),
  mtry = max(1, floor(sqrt(p))),
  sample.fraction = ifelse(n < 500, 0.8, 0.5),
  honesty = TRUE,
  honesty.fraction = 0.5,
  alpha = 0.05,
  sample.weights = df$iptw_weights_stab
)

cat("Model training completed\n")
calib_test <- test_calibration(cf_model)
print(calib_test)

honesty_scenarios <- c(0.3, 0.5, 0.7)
scenario_names <- c("Low(0.3)", "Moderate(0.5)", "High(0.7)")


ite_results <- list()
ate_values <- numeric(length(honesty_scenarios))

for (i in seq_along(honesty_scenarios)) {
  hf <- honesty_scenarios[i]
  cat(sprintf("\n--- Testing Honesty Fraction = %.1f ---\n", hf))
  
  set.seed(123 + i)  
  cf_sens <- causal_forest(
    X = X_modifiers,
    Y = Y_resid,
    W = T,
    num.trees = 4000,                    
    min.node.size = max(5, floor(n * 0.01)),  
    mtry = max(1, floor(sqrt(p))),      
    sample.fraction = ifelse(n < 500, 0.8, 0.5),  
    honesty = TRUE,                      
    honesty.fraction = hf,               
    alpha = 0.05,                        
    sample.weights = df$iptw_weights_stab  
  )
  
  
  pred_sens <- predict(cf_sens, estimate.variance = FALSE)
  ite_sens <- pred_sens$predictions
  

  ate_val <- mean(ite_sens)
  sd_ite <- sd(ite_sens)
  cv_val <- abs(sd_ite / ate_val) * 100 
  

  ite_results[[scenario_names[i]]] <- ite_sens
  ate_values[i] <- ate_val
  
  cat(sprintf("  Mean ATE: %.3f days\n", ate_val))
  cat(sprintf("  SD of ITE: %.3f\n", sd_ite))
  cat(sprintf("  CV (变异系数): %.2f%%\n", cv_val))
  
  
  assign(paste0("cf_hf_", gsub("\\.", "", as.character(hf))), cf_sens)
}


ate_cv <- sd(ate_values) / abs(mean(ate_values)) * 100
cat(sprintf("ATE across honesty fractions: %.3f, %.3f, %.3f\n", 
            ate_values[1], ate_values[2], ate_values[3]))
cat(sprintf("Inter-scenario CV (场景间变异系数): %.2f%%\n", ate_cv))


if (ate_cv < 2.0) {
  cat("✓ 结果稳健：诚实性比例改变对ATE估计影响较小 (CV < 2%)\n")
} else {
  cat("⚠ 提示：诚实性比例对估计结果影响较大，需谨慎解释\n")
}


cat("\nPairwise Correlations (ITE一致性):\n")
cor_03_05 <- cor(ite_results[[1]], ite_results[[2]])
cor_05_07 <- cor(ite_results[[2]], ite_results[[3]])
cor_03_07 <- cor(ite_results[[1]], ite_results[[3]])

cat(sprintf("  Low(0.3) vs Moderate(0.5): r = %.4f\n", cor_03_05))
cat(sprintf("  Moderate(0.5) vs High(0.7): r = %.4f\n", cor_05_07))
cat(sprintf("  Low(0.3) vs High(0.7): r = %.4f\n", cor_03_07))

if (min(cor_03_05, cor_05_07, cor_03_07) > 0.95) {
  cat("✓ ITE估计高度一致 (r > 0.95)\n")
}

sens_df <- data.frame(
  Honesty_Fraction = rep(scenario_names, each = n),
  ITE = c(ite_results[[1]], ite_results[[2]], ite_results[[3]]),
  ATE = rep(ate_values, each = n)
)

p_sens <- ggplot(sens_df, aes(x = ITE, fill = Honesty_Fraction)) +
  geom_density(alpha = 0.5) +
  geom_vline(data = data.frame(Honesty_Fraction = scenario_names, ATE = ate_values),
             aes(xintercept = ATE, color = Honesty_Fraction),
             linetype = "dashed", linewidth = 1) +
  scale_fill_manual(values = c("#E69F00", "#56B4E9", "#009E73")) +
  scale_color_manual(values = c("#E69F00", "#56B4E9", "#009E73")) +
  labs(
    title = "Sensitivity to Honesty Fraction",
    subtitle = sprintf("CV of ATE across scenarios: %.2f%% (Threshold: <2%%)", ate_cv),
    x = "Individual Treatment Effect (days)",
    y = "Density",
    fill = "Honesty Fraction",
    color = "Honesty Fraction"
  ) +
  theme_classic(base_size = 12) +
  theme(legend.position = "top")

ggsave(paste0(save_path, "Fig_S1_Honesty_Fraction_Sensitivity.png"), 
       p_sens, width = 10, height = 6, dpi = 600)
sens_summary <- data.frame(
  Honesty_Fraction = c("0.3 (Low)", "0.5 (Moderate/Main)", "0.7 (High)"),
  Mean_ATE = round(ate_values, 3),
  SD_ITE = round(c(sd(ite_results[[1]]), sd(ite_results[[2]]), sd(ite_results[[3]])), 3),
  CV_Percent = round(c(
    abs(sd(ite_results[[1]])/ate_values[1])*100,
    abs(sd(ite_results[[2]])/ate_values[2])*100,
    abs(sd(ite_results[[3]])/ate_values[3])*100
  ), 2)
)

write.csv(sens_summary, 
          paste0(save_path, "Table_Sensitivity_Honesty_Fraction.csv"), 
          row.names = FALSE)

cat(sprintf("\n结果已保存至: %sTable_Sensitivity_Honesty_Fraction.csv\n", save_path))
set.seed(456)
folds <- sample(rep(1:5, length.out = n))
importance_cv <- matrix(0, nrow = p, ncol = 5)

for (k in 1:5) {
  train_idx <- which(folds != k)
  cf_fold <- causal_forest(
    X = X_modifiers[train_idx, , drop = FALSE],
    Y = Y_resid[train_idx],
    W = T[train_idx],
    num.trees = 2000,
    min.node.size = 5,
    sample.weights = df$iptw_weights_stab[train_idx]
  )
  importance_cv[, k] <- variable_importance(cf_fold)
}

importance_mean <- rowMeans(importance_cv)
importance_se <- apply(importance_cv, 1, sd)
importance_ci_lower <- importance_mean - 1.96 * (importance_se / sqrt(5))
importance_ci_upper <- importance_mean + 1.96 * (importance_se / sqrt(5))

importance_df <- data.frame(
  Variable = colnames(X_modifiers),
  Importance_Mean = importance_mean,
  SE = importance_se / sqrt(5),
  CI_Lower = importance_ci_lower,
  CI_Upper = importance_ci_upper,
  Core_Screening = ifelse(colnames(X_modifiers) %in% core_markers, "Yes", "No"),
  stringsAsFactors = FALSE
) %>% arrange(desc(Importance_Mean))

cat("Top 10 Important Variables:\n")
print(head(importance_df, 10), row.names = FALSE)
marker_idx <- which(colnames(X_modifiers) %in% importance_df$Variable[1:8])
cat("Selected markers for subgroup profiling (Top 8):", 
    paste(colnames(X_modifiers)[marker_idx], collapse = ", "), "\n")

importance_sd <- apply(importance_cv, 1, sd)  
cv_percent <- (importance_sd / importance_mean) * 100


top8_idx <- order(importance_mean, decreasing = TRUE)[1:8]
top8_cv <- cv_percent[top8_idx]

cat("Top 8 CV (%):\n")
print(round(top8_cv, 1))

if(all(top8_cv < 15)) {
  cat("✓ CV < 15%\n")
} else {
  cat("⚠ 实际CV:", round(mean(top8_cv), 1))
}

n_perm <- 100
p <- ncol(X_modifiers)
perm_importance <- matrix(0, nrow = p, ncol = n_perm)

set.seed(789)  
for (i in 1:n_perm) {
  T_perm <- sample(df$treatment)  
  cf_perm <- causal_forest(X_modifiers, Y_resid, T_perm, num.trees = 500, 
                           sample.weights = df$iptw_weights_stab)
  perm_importance[, i] <- variable_importance(cf_perm)
  if(i %% 20 == 0) cat("Permutation", i, "/", n_perm, "completed\n")
}

p_values <- sapply(1:nrow(importance_df), function(j) {
  idx <- which(colnames(X_modifiers) == importance_df$Variable[j])
  mean(perm_importance[idx, ] >= importance_df$Importance_Mean[j])
})
importance_df$P_Value <- p_values
importance_df$P_Adj_FDR <- p.adjust(importance_df$P_Value, method = "fdr")


importance_df <- importance_df %>%
  mutate(Significance = case_when(
    P_Adj_FDR < 0.10 & CI_Lower > 0 ~ "Primary (FDR<0.10)",
    P_Value < 0.05 & CI_Lower > 0 ~ "Secondary (P<0.05)",
    TRUE ~ "NS"
  ))

core_primary <- importance_df %>% filter(Significance == "Primary (FDR<0.10)")
core_secondary <- importance_df %>% filter(Significance == "Secondary (P<0.05)")

cat("核心标志物 (FDR<0.10):", nrow(core_primary), "\n")
if(nrow(core_primary) > 0) cat(paste(core_primary$Variable, collapse=", "), "\n")
cat("候选标志物 (P<0.05):", nrow(core_secondary), "\n")

pred_cf <- predict(cf_model, estimate.variance = TRUE)
ite <- pred_cf$predictions
ite_se <- sqrt(pred_cf$variance.estimates)

df$ITE <- ite
df$ITE_lower <- ite - 1.96 * ite_se
df$ITE_upper <- ite + 1.96 * ite_se

cat(paste("ITE range:", round(min(ite), 2), "to", round(max(ite), 2), "days\n"))
cat(paste("Mean ITE:", round(mean(ite), 2), "days\n"))

cat("\n=== Step 17: E-value Calculation ===\n")

if (hr < 1) {
  observed_rr <- 1/hr
  ci_lower_rr <- 1/ci_upper  
  ci_upper_rr <- 1/ci_lower
} else {
  observed_rr <- hr
  ci_lower_rr <- ci_lower
  ci_upper_rr <- ci_upper
}

e_value <- observed_rr + sqrt(observed_rr * (observed_rr - 1))
e_value_lower <- ci_lower_rr + sqrt(ci_lower_rr * (ci_lower_rr - 1))

cat("Observed HR:", round(hr, 2), "\n")
cat("Observed RR for E-value:", round(observed_rr, 2), "\n")
cat("E-value:", round(e_value, 2), "\n")
cat("E-value (lower CI):", round(e_value_lower, 2), "\n")
cat("Interpretation: Unmeasured confounder needs RR of", round(e_value, 2), 
    "to fully explain away the observed effect.\n")
calc_weighted_smd <- function(x1, x2, w1, w2) {
  mu1 <- weighted.mean(x1, w1, na.rm = TRUE)
  mu2 <- weighted.mean(x2, w2, na.rm = TRUE)
  var1 <- Hmisc::wtd.var(x1, w1, na.rm = TRUE)
  var2 <- Hmisc::wtd.var(x2, w2, na.rm = TRUE)
  pooled_sd <- sqrt((var1 + var2) / 2)
  if (pooled_sd == 0 || is.na(pooled_sd)) return(0)
  return((mu1 - mu2) / pooled_sd)
}


ite_tertiles <- quantile(df$ITE, probs = c(1/3, 2/3), na.rm = TRUE)



ite_tertiles <- quantile(df$ITE, probs = c(1/3, 2/3), na.rm = TRUE)
cat("ITE Tertiles: T1 ≤", round(ite_tertiles[1], 2), ", T2 ≤", round(ite_tertiles[2], 2), "\n")

df$benefit_cat <- case_when(
  df$ITE <= ite_tertiles[1] ~ "High Benefit (Q1)",        
  df$ITE <= ite_tertiles[2] ~ "Moderate Benefit (Q2)",     
  TRUE ~ "Low Benefit (Q3)"                                
)

cat("Benefit categories (Tertile-based):\n"); print(table(df$benefit_cat))

tertile_profiles <- list()

for (tertile in c("High Benefit (Q1)", "Moderate Benefit (Q2)", "Low Benefit (Q3)")) {
  in_grp <- which(df$benefit_cat == tertile)
  out_grp <- which(df$benefit_cat != tertile)
  
  
  smd_vals <- sapply(marker_idx, function(idx) {
    calc_weighted_smd(X_modifiers[in_grp, idx], X_modifiers[out_grp, idx],
                      df$iptw_weights_stab[in_grp], df$iptw_weights_stab[out_grp])
  })
  
  cat(paste0("\n[", tertile, "] (n=", length(in_grp), 
             ", Mean ITE=", round(mean(df$ITE[in_grp]), 1), " days):\n"))
  
  prof <- data.frame(
    Method = "Tertile-based",
    Subgroup = tertile,
    Variable = colnames(X_modifiers)[marker_idx],
    SMD = round(smd_vals, 2),
    Within_Group_Mean = round(sapply(marker_idx, function(idx) 
      weighted.mean(X_modifiers[in_grp, idx], df$iptw_weights_stab[in_grp], na.rm = TRUE)), 2),
    Overall_Mean = round(sapply(marker_idx, function(idx) 
      weighted.mean(X_modifiers[, idx], df$iptw_weights_stab, na.rm = TRUE)), 2),
    Direction = ifelse(smd_vals > 0, "Higher than others", "Lower than others")
  ) %>% arrange(desc(abs(SMD)))
  
  print(prof, row.names = FALSE)
  tertile_profiles[[tertile]] <- prof
}


set.seed(123)
km_res <- kmeans(df$ITE, centers = 3, nstart = 25, algorithm = "Lloyd")
clust_order <- order(aggregate(ITE ~ km_res$cluster, data = df, FUN = mean)$ITE)
df$subgroup <- factor(km_res$cluster, levels = clust_order,
                      labels = c("High Responder", "Moderate Responder", "Low Responder"))

cat("Cluster subgroups:\n"); print(table(df$subgroup))

subgroup_profiles <- list()

for (i in 1:3) {
  raw_clust <- clust_order[i]
  sg_name <- levels(df$subgroup)[i]
  in_grp <- which(km_res$cluster == raw_clust)
  out_grp <- which(km_res$cluster != raw_clust)
  
  
  smd_vals <- sapply(marker_idx, function(idx) {
    calc_weighted_smd(X_modifiers[in_grp, idx], X_modifiers[out_grp, idx],
                      df$iptw_weights_stab[in_grp], df$iptw_weights_stab[out_grp])
  })
  
  cat(paste0("\n[", sg_name, "] (n=", length(in_grp), 
             ", Mean ITE=", round(mean(df$ITE[in_grp]), 1), " days):\n"))
  
  prof <- data.frame(
    Method = "K-means",
    Subgroup = sg_name,
    Variable = colnames(X_modifiers)[marker_idx],
    SMD = round(smd_vals, 2),
    Within_Group_Mean = round(sapply(marker_idx, function(idx) 
      weighted.mean(X_modifiers[in_grp, idx], df$iptw_weights_stab[in_grp], na.rm = TRUE)), 2),
    Overall_Mean = round(sapply(marker_idx, function(idx) 
      weighted.mean(X_modifiers[, idx], df$iptw_weights_stab, na.rm = TRUE)), 2),
    Direction = ifelse(smd_vals > 0, "Higher than others", "Lower than others")
  ) %>% arrange(desc(abs(SMD)))
  
  print(prof, row.names = FALSE)
  subgroup_profiles[[sg_name]] <- prof
}

high_tertile <- tertile_profiles[["High Benefit (Q1)"]]$SMD
high_kmeans <- subgroup_profiles[["High Responder"]]$SMD
correlation <- cor(abs(high_tertile), abs(high_kmeans))

cat(paste0("Correlation of SMD profiles (High Benefit vs High Responder): ", 
           round(correlation, 3), "\n"))
if(correlation > 0.8) {
  cat("✓ High consistency between methods: Biological profile is robust to subgroup definition.\n")
} else {
  cat("⚠ Moderate consistency: Check for algorithm-specific artifacts.\n")
}


if (length(tertile_profiles) > 0 & length(subgroup_profiles) > 0) {
  prof_tertile_all <- do.call(rbind, tertile_profiles)
  prof_kmeans_all <- do.call(rbind, subgroup_profiles)
  prof_combined <- rbind(prof_tertile_all, prof_kmeans_all)
  
  write.csv(prof_combined, 
            paste0(save_path, "Results_Subgroup_Profiles_Both_Methods.csv"), 
            row.names = FALSE)
  cat("✓ Combined profiles saved: Tertile-based and K-means methods\n")
}


write.csv(do.call(rbind, tertile_profiles), 
          paste0(save_path, "Results_Tertile_Profiles.csv"), 
          row.names = FALSE)


write.csv(do.call(rbind, subgroup_profiles), 
          paste0(save_path, "Results_Kmeans_Profiles.csv"), 
          row.names = FALSE)


cat(paste0("\nClinical Reference Summary:\n"))
cat(paste0("Tertile-based High Benefit (Q1): ≤", round(ite_tertiles[1], 2), " days, n=", 
           sum(df$benefit_cat == "High Benefit (Q1)"), "\n"))
cat(paste0("K-means High Responder: Mean ITE=", 
           round(mean(df$ITE[df$subgroup == "High Responder"]), 2), " days, n=", 
           sum(df$subgroup == "High Responder"), "\n"))
comparison_data_tertile <- data.frame(Variable = colnames(X_modifiers)[marker_idx], stringsAsFactors = FALSE)


for (sg in c("High Benefit (Q1)", "Moderate Benefit (Q2)", "Low Benefit (Q3)")) {
  in_grp <- which(df$benefit_cat == sg)
  out_grp <- which(df$benefit_cat != sg)
  comparison_data_tertile[[sg]] <- sapply(marker_idx, function(idx) {
    round(calc_weighted_smd(X_modifiers[in_grp, idx], X_modifiers[out_grp, idx],
                            df$iptw_weights_stab[in_grp], df$iptw_weights_stab[out_grp]), 2)
  })
}

write.csv(comparison_data_tertile, 
          paste0(save_path, "Results_Tertile_Heatmap_Data.csv"), 
          row.names = FALSE)

cat("\nStep 18 completed: Both Tertile-based and K-means profiling finished.\n")
p1 <- ggplot(ps_df, aes(x = PS, fill = Treatment)) +
  geom_density(alpha = 0.6) +
  geom_vline(xintercept = c(0.05, 0.95), linetype = "dashed", color = "red") +
  scale_fill_manual(values = c("#E69F00", "#56B4E9")) +
  labs(title = "Propensity Score Distribution", x = "PS", y = "Density") +
  theme_classic()
ggsave(paste0(save_path, "Fig1_PS_Distribution.png"), p1, width = 8, height = 6, dpi = 300)

df$benefit_cat <- factor(df$benefit_cat, 
                         levels = c("High Benefit (Q1)", "Moderate Benefit (Q2)", "Low Benefit (Q3)"))
p2 <- ggplot(df, aes(x = ITE, fill = benefit_cat)) +
  geom_density(alpha = 0.6) +
  geom_vline(xintercept = 0, linetype = "dashed", color = "red") +
  scale_fill_manual(values = c("High Benefit (Q1)" = "#2E8B57", "Moderate Benefit (Q2)" = "#90EE90",
                               "Harm (Q3)" = "#DC143C")) +
  labs(title = "ITE Distribution", x = "Treatment Effect (days)", y = "Density") +
  theme_classic()
ggsave(paste0(save_path, "Fig2_ITE_Distribution.png"), p2, width = 10, height = 6, dpi = 600)

p3 <- ggplot(df, aes(x = subgroup, y = ITE, fill = subgroup)) +
  geom_boxplot(alpha = 0.7) + geom_jitter(width = 0.2, alpha = 0.3, size = 1) +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red") +
  geom_hline(yintercept = mean(df$ITE), linetype = "dotted", color = "blue", alpha = 0.7) +
  scale_fill_manual(values = c("High Responder" = "#2E8B57", "Moderate Responder" = "#E69F00", "Low Responder" = "#56B4E9")) +
  labs(title = "ITE by Subgroup", x = "Subgroup", y = "Days") + 
  theme_classic() + theme(legend.position = "none")
ggsave(paste0(save_path, "Fig3_ITE_Subgroups.png"), p3, width = 8, height = 6, dpi = 600)

imp_plot <- importance_df %>% head(15) %>% mutate(
  Significance = case_when(P_Adj_FDR < 0.05 ~ "***", P_Adj_FDR < 0.10 ~ "*", P_Value < 0.05 ~ ".", TRUE ~ "ns"),
  Label = paste0(Variable, " ", Significance)
)
p4 <- ggplot(imp_plot, aes(x = reorder(Label, Importance_Mean), y = Importance_Mean, fill = Core_Screening)) +
  geom_col(alpha = 0.8) +
  geom_errorbar(aes(ymin = CI_Lower, ymax = CI_Upper), width = 0.2, color = "#003366") +
  coord_flip() +
  scale_fill_manual(values = c("Yes" = "#D55E00", "No" = "#6699CC")) +
  labs(title = "Single-Variable Importance", x = "Marker", y = "Importance") + 
  theme_classic()
ggsave(paste0(save_path, "Fig4_SingleVariable_Importance.png"), p4, width = 10, height = 8, dpi = 600)


heatmap_long <- comparison_data_tertile %>%
  pivot_longer(-Variable, names_to = "Subgroup", values_to = "SMD") %>%
  mutate(Subgroup = factor(Subgroup, 
                           levels = c("High Benefit (Q1)", "Moderate Benefit (Q2)", "Low Benefit (Q3)"),
                           labels = c("High Responder", "Moderate Responder", "Low Responder")),
         SMD_plot = pmax(pmin(SMD, 2.5), -2.5),
         Label = sprintf("%.2f", SMD),
         Significance = case_when(abs(SMD) > 0.8 ~ "large", abs(SMD) > 0.5 ~ "moderate", 
                                  abs(SMD) > 0.2 ~ "small", TRUE ~ "negligible"))

p5 <- ggplot(heatmap_long, aes(x = Subgroup, y = reorder(Variable, abs(SMD)), fill = SMD_plot)) +
  geom_tile(color = "white", linewidth = 0.5) +
  geom_text(aes(label = Label, color = Significance), size = 3.5, fontface = "bold") +
  scale_fill_gradient2(low = "#D73027", mid = "#FFFFBF", high = "#1A9850", 
                       midpoint = 0, limits = c(-2.5, 2.5), name = "SMD") +
  scale_color_manual(values = c("large" = "white", "moderate" = "black", 
                                "small" = "black", "negligible" = "grey50"), guide = "none") +
  labs(title = "Multi-Variable Pattern (Weighted SMD)", x = "Subgroup", y = "Marker") +
  theme_minimal() + 
  theme(axis.text.x = element_text(face = "bold", size = 11), 
        axis.text.y = element_text(face = "bold", size = 10),
        axis.title = element_text(face = "bold", size = 12),
        panel.grid = element_blank(),
        plot.title = element_text(face = "bold", size = 14))
ggsave(paste0(save_path, "Fig5_MultiVariable_Heatmap.png"), p5, width = 10, height = 8, dpi = 600)

forest_data <- data.frame(
  Subgroup = c("High Responder", "Moderate Responder", "Low Responder", "Overall"),
  ITE_Mean = c(
    weighted.mean(df$ITE[df$subgroup == "High Responder"], df$iptw_weights_stab[df$subgroup == "High Responder"]),
    weighted.mean(df$ITE[df$subgroup == "Moderate Responder"], df$iptw_weights_stab[df$subgroup == "Moderate Responder"]),
    weighted.mean(df$ITE[df$subgroup == "Low Responder"], df$iptw_weights_stab[df$subgroup == "Low Responder"]),
    weighted.mean(df$ITE, df$iptw_weights_stab)
  ),
  N = c(sum(df$subgroup == "High Responder"), sum(df$subgroup == "Moderate Responder"),
        sum(df$subgroup == "Low Responder"), nrow(df))
) %>% mutate(
  ITE_SE = c(
    sqrt(wtd.var(df$ITE[df$subgroup == "High Responder"], df$iptw_weights_stab[df$subgroup == "High Responder"]) / N[1]),
    sqrt(wtd.var(df$ITE[df$subgroup == "Moderate Responder"], df$iptw_weights_stab[df$subgroup == "Moderate Responder"]) / N[2]),
    sqrt(wtd.var(df$ITE[df$subgroup == "Low Responder"], df$iptw_weights_stab[df$subgroup == "Low Responder"]) / N[3]),
    sqrt(wtd.var(df$ITE, df$iptw_weights_stab) / N[4])
  ),
  CI_Lower = ITE_Mean - 1.96 * ITE_SE,
  CI_Upper = ITE_Mean + 1.96 * ITE_SE
)

p6 <- ggplot(forest_data, aes(x = ITE_Mean, y = reorder(Subgroup, ITE_Mean), color = Subgroup)) +
  geom_point(aes(size = N), alpha = 0.8) +
  geom_errorbarh(aes(xmin = CI_Lower, xmax = CI_Upper), height = 0.2, linewidth = 1) +
  geom_vline(xintercept = 0, linetype = "dashed", color = "red") +
  scale_color_manual(values = c("High Responder" = "#2E8B57", "Moderate Responder" = "#E69F00", 
                                "Low Responder" = "#56B4E9", "Overall" = "black")) +
  scale_size_continuous(range = c(3, 6)) +
  labs(title = "Forest Plot: Treatment Effect by Subgroup", x = "ITE (days)", y = "") +
  theme_classic() + theme(legend.position = "none")
ggsave(paste0(save_path, "Fig6_Forest_Plot.png"), p6, width = 10, height = 6, dpi = 600)

cat("All 6 visualizations completed\n")


write.csv(df %>% select(treatment, remission_time, time, event, ITE, ITE_lower, ITE_upper, 
                        subgroup, benefit_cat, iptw_weights_stab, 
                        all_of(available_lab_vars)),  
          paste0(save_path, "Results_Individual_ITE.csv"), 
          row.names = FALSE)

write.csv(importance_df, paste0(save_path, "Results_Feature_Importance.csv"), row.names = FALSE)

if (length(subgroup_profiles) > 0) {
  prof_all <- do.call(rbind, lapply(names(subgroup_profiles), function(n) 
    cbind(Subgroup = n, subgroup_profiles[[n]])))
  write.csv(prof_all, paste0(save_path, "Results_Subgroup_Profiles.csv"), row.names = FALSE)
}


hr_logit_safe <- if(exists("cox_logit")) exp(coef(cox_logit)) else NA
ci_logit_safe <- if(exists("cox_logit")) exp(confint(cox_logit)) else c(NA, NA)
hr_ato_safe <- if(exists("cox_ato")) exp(coef(cox_ato)) else NA
ci_ato_safe <- if(exists("cox_ato")) exp(confint(cox_ato)) else c(NA, NA)
hr_trimmed_safe <- if(exists("cox_trimmed")) exp(coef(cox_trimmed)) else NA
ci_trimmed_safe <- if(exists("cox_trimmed")) exp(confint(cox_trimmed)) else c(NA, NA)

summary_final <- data.frame(
  Metric = c("Total Sample", "Yindan Group", "Control Group", "Events", "Censored", 
             "Weight Truncation", "HR (IPTW)", "P-value", "E-value", "E-value (lower CI)",
             "High Responder %", "Expected Benefit (days)"),
  Value = c(nrow(df), 
            sum(df$treatment), 
            sum(1-df$treatment), 
            sum(df$event), 
            sum(1-df$event),
            "1%-99%", 
            paste0(round(hr, 2), " (", round(ci_lower, 2), "-", round(ci_upper, 2), ")"),
            format.pval(p_val, eps = 0.001),
            round(e_value, 2),
            round(e_value_lower, 2),
            paste0(round(mean(df$subgroup == "High Responder", na.rm = TRUE)*100, 1), "%"),
            round(abs(mean(df$ITE[df$subgroup == "High Responder"], na.rm = TRUE)), 1))
)
write.csv(summary_final, paste0(save_path, "Results_Summary.csv"), row.names = FALSE)


if (exists("cox_logit") && exists("cox_ato") && exists("cox_trimmed")) {
  sensitivity_summary <- data.frame(
    Method = c("Main (GBM-IPTW)", "Logistic PS", "Overlap Weighting", "Trimmed (Positivity)"),
    HR = c(hr, hr_logit_safe, hr_ato_safe, hr_trimmed_safe),
    CI_Lower = c(ci_lower, ci_logit_safe[1], ci_ato_safe[1], ci_trimmed_safe[1]),
    CI_Upper = c(ci_upper, ci_logit_safe[2], ci_ato_safe[2], ci_trimmed_safe[2])
  )
  write.csv(sensitivity_summary, paste0(save_path, "Sensitivity_Summary.csv"), row.names = FALSE)
  cat("Sensitivity analysis completed and saved\n")
} else {
  cat("Warning: Some sensitivity analyses skipped (missing model objects)\n")
}


write.csv(screen_df, 
          paste0(save_path, "Table_S1_Screening_Importance_All.csv"), 
          row.names = FALSE)


write.csv(screen_df %>% head(10) %>% mutate(Rank = row_number()), 
          paste0(save_path, "Table_S2_Core_Markers_Top10.csv"), 
          row.names = FALSE)

cat("Screening results saved: All variables (n=", nrow(screen_df), ") and Top 10\n")
cat("\n=== Analysis Completed ===\n")


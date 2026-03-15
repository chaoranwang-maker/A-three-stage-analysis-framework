
packages <- c("tidyverse", "ggplot2", "readxl", "stringr", "mice", 
              "cobalt", "survival", "survminer", "WeightIt", 
              "Hmisc", "gridExtra")
installed <- rownames(installed.packages())
for (pkg in packages) {
  if (!pkg %in% installed) install.packages(pkg)
}
library(tidyverse)
library(ggplot2)
library(readxl)
library(stringr)
library(mice)
library(cobalt)
library(survival)
library(survminer)
library(WeightIt)
library(Hmisc)
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
  
  
  Clinical_classification = as.numeric(`出院分型`),
  
  
  treatment = as.numeric(`是否用药`),
  
  
  anti_viral = as.numeric(`是否联合抗病毒`),
  
  
  admission_temp = as.numeric(`入院体温`)
)


df$vaccine_num[is.na(df$vaccine_num)] <- 10


df$vaccine <- factor(df$vaccine_num, 
                     levels = c(0,1,2,3,10),
                     labels = c("Unvaccinated", "Dose1", "Dose2", "Dose3plus", "Unknown"))


df$vaccine_3plus <- ifelse(df$vaccine_num %in% c(3, 4), 1, 0)

cat("Variable check:\n")
cat("Age range:", range(df$age, na.rm = TRUE), "\n")
cat("Sex (numeric):", class(df$sex), "Values:", unique(df$sex), "\n")
cat("Anti_viral (numeric):", class(df$anti_viral), "Values:", unique(df$anti_viral), "\n")
cat("Vaccine (factor levels):", levels(df$vaccine), "\n")
cat("Clinical_classification (numeric):", class(df$Clinical_classification), "Values:", unique(df$Clinical_classification), "\n")


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


basic_vars <- c("age", "sex", "CCI_score", "smoke", "drink", 
                "Clinical_classification", "anti_viral", "admission_temp",
                "remission_time", "event")

cat("Basic variable missing check:\n")
for (var in basic_vars) {
  n_missing <- sum(is.na(df[[var]]))
  cat(sprintf("  %s: %d missing (%.2f%%)\n", var, n_missing, 100*n_missing/nrow(df)))
  if (n_missing > 0 && var %in% c("age", "sex", "remission_time", "event")) {
    stop(paste("Error: Critical variable", var, "has missing values!"))
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

core_markers <- c("IL.6", "LDH", "HBDH", "ALB", "TP", "LYM.", "TT", "CHE", "DD", "FDP")
cat("Core markers (from Phase 1 screening):\n")
cat(paste(core_markers, collapse = ", "), "\n")

cat("\n=== Step 6: Exact Replication of Phase 1 IPTW + Interaction Cox ===\n")


weighting_vars <- c("age", "sex", "CCI_score", "smoke", "drink",
                    "vaccine", "Clinical_classification", "anti_viral", "admission_temp",
                    core_markers)

weighting_formula <- as.formula(
  paste("treatment ~", paste(intersect(weighting_vars, colnames(df)), collapse = " + "))
)

cat("IPTW formula (exactly as Phase 1 Step 8):\n")
print(weighting_formula)

set.seed(456)
W.out <- weightit(weighting_formula,
                  data = df,
                  method = "gbm",
                  estimand = "ATE",
                  stop.method = "es.mean")

df$iptw_weights <- W.out$weights
weight_q <- quantile(df$iptw_weights, probs = c(0.01, 0.99), na.rm = TRUE)
df$iptw_weights_stab <- pmin(pmax(df$iptw_weights, weight_q[1]), weight_q[2])

cat(sprintf("Weight range after truncation: %.3f - %.3f\n", 
            min(df$iptw_weights_stab), max(df$iptw_weights_stab)))


cox_interact <- coxph(
  Surv(time, event) ~ treatment * anti_viral + treatment * vaccine_3plus,
  data = df,
  weights = iptw_weights_stab,
  robust = TRUE
)

coefs <- summary(cox_interact)$coefficients

cat("\nInteraction P-values:\n")
cat(sprintf("  Treatment x Antiviral: %.4f\n", coefs["treatment:anti_viral", "Pr(>|z|)"]))
cat(sprintf("  Treatment x Vaccine: %.4f\n", coefs["treatment:vaccine_3plus", "Pr(>|z|)"]))


event_table <- df %>%
  group_by(anti_viral, vaccine_3plus) %>%
  summarise(
    N = n(),
    Events = sum(event),
    Median_FU = median(time),
    .groups = 'drop'
  ) %>%
  mutate(
    Subgroup_Label = paste0(
      ifelse(anti_viral == 0, "No Antiviral", "Antiviral"),
      " + ",
      ifelse(vaccine_3plus == 0, "<3 Doses", "3+ Doses")
    )
  )

cat("\nSubgroup Descriptive Statistics:\n")
print(event_table[, c("Subgroup_Label", "N", "Events", "Median_FU")])


beta_treat <- coefs["treatment", "coef"]
beta_antiviral <- coefs["treatment:anti_viral", "coef"]
beta_vaccine <- coefs["treatment:vaccine_3plus", "coef"]

se_treat <- coefs["treatment", "se(coef)"]
se_antiviral <- coefs["treatment:anti_viral", "se(coef)"]
se_vaccine <- coefs["treatment:vaccine_3plus", "se(coef)"]


subgroups <- expand.grid(
  anti_viral = c(0, 1),
  vaccine_3plus = c(0, 1)
) %>%
  mutate(
    log_hr = beta_treat + anti_viral * beta_antiviral + vaccine_3plus * beta_vaccine,
    HR = exp(log_hr),
    var_log_hr = se_treat^2 + 
      (anti_viral^2) * se_antiviral^2 + 
      (vaccine_3plus^2) * se_vaccine^2,
    HR_lower = exp(log_hr - 1.96 * sqrt(var_log_hr)),
    HR_upper = exp(log_hr + 1.96 * sqrt(var_log_hr)),
    Group = paste0(ifelse(anti_viral == 0, "No Antiviral", "Antiviral"),
                   " + ",
                   ifelse(vaccine_3plus == 0, "<3 Doses", "3+ Doses")),
    Antiviral_Factor = ifelse(anti_viral == 0, "No Antiviral", "Antiviral")
  ) %>%
  left_join(event_table, by = c("anti_viral", "vaccine_3plus"))

print(subgroups[, c("Group", "HR", "HR_lower", "HR_upper", "N", "Events")])



ref_data <- data.frame(
  Group = "Control (Reference)", 
  HR = 1, 
  HR_lower = 1, 
  HR_upper = 1, 
  Antiviral_Factor = "No Antiviral",
  N = event_table$N[event_table$anti_viral==0 & event_table$vaccine_3plus==0],
  Events = event_table$Events[event_table$anti_viral==0 & event_table$vaccine_3plus==0]
)

plot_data <- bind_rows(ref_data, subgroups)
plot_data$Label <- with(plot_data, paste0(Group, " (N=", N, ", Events=", round(Events, 0), ")"))

p_interact_antiviral <- coefs["treatment:anti_viral", "Pr(>|z|)"]
p_interact_vaccine <- coefs["treatment:vaccine_3plus", "Pr(>|z|)"]

p_forest <- ggplot(plot_data, aes(x = HR, y = reorder(Label, desc(Label)), color = Antiviral_Factor)) +
  geom_point(aes(size = ifelse(Group == "Control (Reference)", 3, 5)), show.legend = c(size = FALSE, color = TRUE)) +
  geom_errorbarh(aes(xmin = HR_lower, xmax = HR_upper), height = 0.2, linewidth = 0.8) +
  geom_vline(xintercept = 1, linetype = "dashed", color = "red", alpha = 0.5) +
  scale_x_log10(limits = c(0.3, 3), breaks = c(0.3, 0.5, 0.7, 1, 1.5, 2, 3)) +
  scale_color_manual(values = c("No Antiviral" = "#0072B5", "Antiviral" = "#D55E00")) +
  labs(
    x = "Hazard Ratio (log scale)",
    y = "",
    title = "Effect Modification by Antiviral and Vaccine Status",
    subtitle = paste0("Single imputed dataset (action=1) matching Phase 1 standard\n",
                      "Antiviral×Treatment P=", format(round(p_interact_antiviral, 4), nsmall=4),
                      "; Vaccine×Treatment P=", format(round(p_interact_vaccine, 4), nsmall=4)),
    caption = "IPTW: Exact replication of Phase 1 (GBM-IPTW with vaccine as 5-level factor)"
  ) +
  theme_classic(base_size = 12) +
  theme(plot.title = element_text(face = "bold", size = 14))

print(p_forest)
ggsave(paste0(save_path, "Interaction_Forest_Exact_Phase1.png"), 
       p_forest, width = 10, height = 6, dpi = 600)


write.csv(data.frame(
  Variable = rownames(coefs),
  Coef = coefs[, "coef"],
  HR = exp(coefs[, "coef"]),
  SE = coefs[, "se(coef)"],
  P_value = coefs[, "Pr(>|z|)"]
), paste0(save_path, "Interaction_Cox_Results.csv"), row.names = FALSE)

write.csv(subgroups, paste0(save_path, "Subgroup_HR_Results.csv"), row.names = FALSE)
write.csv(event_table, paste0(save_path, "Subgroup_Events.csv"), row.names = FALSE)

cat("\n=== Analysis Completed ===\n")
cat("Method: Single imputation (action=1), exact match to Phase 1\n")
cat(sprintf("Interaction P-values: Antiviral=%.4f, Vaccine=%.4f\n", 
            p_interact_antiviral, p_interact_vaccine))


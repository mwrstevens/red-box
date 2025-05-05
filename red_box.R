# -- Green Box Red Box - Data processing and analysis --
# This script analyses the green box data with proper package installation
# and portable directory handling

# -- Package Installation and Loading --
required_packages <- c("dplyr", "readxl", "ppcor", "pROC")

# Check if packages are installed, install if necessary
for (pkg in required_packages) {
  if (!require(pkg, character.only = TRUE, quietly = TRUE)) {
    message(paste("Installing package:", pkg))
    install.packages(pkg, repos = "https://cran.r-project.org")
    library(pkg, character.only = TRUE)
  } else {
    message(paste("Package", pkg, "is already installed and loaded"))
  }
}

# -- Setup Project Directory --
# Create a function to set up the working directory in a portable way
setup_project <- function(data_file = "red_box.xlsx", user_dir = NULL) {
  # Use current directory if no directory is specified
  if (is.null(user_dir)) {
    main_dir <- getwd()
    message("Using current working directory: ", main_dir)
  } else {
    if (dir.exists(user_dir)) {
      main_dir <- user_dir
      setwd(main_dir)
      message("Working directory set to: ", main_dir)
    } else {
      stop("Directory does not exist: ", user_dir)
    }
  }
  
  # Create project structure
  project_dir <- file.path(main_dir, "Green box")
  if (!dir.exists(project_dir)) dir.create(project_dir)
  
  subdirs <- c("Manuscript Files", "Supplementary Files")
  dir_paths <- list()
  
  for (subdir in subdirs) {
    path <- file.path(project_dir, subdir)
    if (!dir.exists(path)) dir.create(path)
    dir_paths[[gsub(" ", "_", tolower(subdir))]] <- path
  }
  
  # Check if data file exists in the project directory
  data_path <- file.path(project_dir, data_file)
  if (!file.exists(data_path)) {
    # Warn user that data file is not found
    warning(paste("Data file", data_file, "not found in project directory:", data_path))
    warning("Please specify data_file parameter with correct path or place the file in the appropriate directory.")
    # Create an empty dummy file for testing purposes
    message("Creating an empty dummy file for testing purposes. Replace with actual data file.")
    write.csv(data.frame(), data_path)
  }
  
  return(list(
    main_dir = main_dir,
    project_dir = project_dir,
    dir_paths = dir_paths,
    data_path = data_path
  ))
}

# -- Data Processing --
process_data <- function(df) {
  # Convert categorical variables to factors and create numeric versions
  categorical_vars <- list(
    nationality = "nationality_recoded",
    employment = "employment_recoded",
    education = "education_recoded"
  )
  
  for (var_name in names(categorical_vars)) {
    df[[var_name]] <- as.factor(df[[var_name]])
    df[[categorical_vars[[var_name]]]] <- as.numeric(df[[var_name]])
  }
  
  # Create binary status variables
  df$IGD_status <- as.factor(df$IGD_status)
  df$IGD_status_binary <- as.numeric(df$IGD_status == "IGD")
  df$GD_status <- as.factor(df$GD_status)
  df$GD_status_binary <- as.numeric(df$GD_status == "GD")
  
  # Calculate combined variables
  df$green_plus_red <- df$green_box + df$red_box
  df$red_proportion <- (df$red_box / df$green_plus_red) * 100
  
  # Handle missing data
  df$red_box[is.na(df$red_box)] <- median(df$red_box, na.rm = TRUE)
  
  return(df)
}

# -- Descriptive Statistics --
create_descriptive_table <- function(df, group_var = "GD_status", group_value = "GD", table_title = "Sample Characteristics by GD Status") {
  # Calculate total sample size and group sizes
  total_n <- nrow(df)
  non_problem_n <- sum(df[[group_var]] != group_value, na.rm = TRUE)
  problem_n <- sum(df[[group_var]] == group_value, na.rm = TRUE)
  
  # Function to format means and SDs
  format_mean_sd <- function(x) {
    mean_val <- round(mean(x, na.rm = TRUE), 1)
    sd_val <- round(sd(x, na.rm = TRUE), 1)
    return(paste0(mean_val, " (", sd_val, ")"))
  }
  
  # Function to format counts and percentages
  format_count_pct <- function(x, n) {
    count <- sum(x, na.rm = TRUE)
    pct <- round(count/n * 100, 1)
    return(paste0(count, " (", pct, "%)"))
  }
  
  # Function to perform t-test between groups
  perform_ttest <- function(x, group) {
    if(sum(!is.na(x[group == group_value])) < 2 || sum(!is.na(x[group != group_value])) < 2) {
      return(list(t_value = "N/A", p_value = "N/A"))
    }
    
    t_result <- tryCatch({
      test <- t.test(x ~ group)
      list(
        t_value = round(test$statistic, 1),
        p_value = if(test$p.value < 0.001) "<.001" else round(test$p.value, 3)
      )
    }, error = function(e) {
      list(t_value = "N/A", p_value = "N/A")
    })
    
    return(t_result)
  }
  
  # Function to perform chi-square test for categorical variables
  perform_chisq <- function(x, group) {
    tab <- table(x, group)
    if(any(tab < 5)) {
      # Use Fisher's exact test for small sample sizes
      test <- tryCatch({
        fisher.test(tab)
      }, error = function(e) {
        return(list(p.value = NA))
      })
    } else {
      test <- tryCatch({
        chisq.test(tab)
      }, error = function(e) {
        return(list(p.value = NA))
      })
    }
    
    p_value <- if(is.na(test$p.value)) {
      "N/A"
    } else if(test$p.value < 0.001) {
      "<.001"
    } else {
      round(test$p.value, 3)
    }
    
    return(list(t_value = "N/A", p_value = p_value))
  }
  
  # Initialize the results table
  table_data <- data.frame(
    Characteristic = character(),
    Total = character(),
    NonProblem = character(),
    Problem = character(),
    t = character(),
    p = character(),
    stringsAsFactors = FALSE
  )
  
  # Add age row
  age_ttest <- perform_ttest(df$age, df[[group_var]])
  table_data <- rbind(table_data, data.frame(
    Characteristic = "Age, years; mean (SD)",
    Total = format_mean_sd(df$age),
    NonProblem = format_mean_sd(df$age[df[[group_var]] != group_value]),
    Problem = format_mean_sd(df$age[df[[group_var]] == group_value]),
    t = age_ttest$t_value,
    p = age_ttest$p_value,
    stringsAsFactors = FALSE
  ))
  
  # Add nationality
  nationality_test <- perform_chisq(df$nationality, df[[group_var]])
  table_data <- rbind(table_data, data.frame(
    Characteristic = "Nationality",
    Total = "--",
    NonProblem = "--",
    Problem = "--",
    t = "N/A",
    p = nationality_test$p_value,
    stringsAsFactors = FALSE
  ))
  
  # Get all unique nationalities and add them as subrows
  nationalities <- sort(unique(df$nationality))
  for(nat in nationalities) {
    if(!is.na(nat) && nat != "") {
      is_nat <- df$nationality == nat
      table_data <- rbind(table_data, data.frame(
        Characteristic = paste("  ", nat),
        Total = format_count_pct(is_nat, total_n),
        NonProblem = format_count_pct(is_nat & df[[group_var]] != group_value, non_problem_n),
        Problem = format_count_pct(is_nat & df[[group_var]] == group_value, problem_n),
        t = "-",
        p = "-",
        stringsAsFactors = FALSE
      ))
    }
  }
  
  # Add employment status
  employment_test <- perform_chisq(df$employment, df[[group_var]])
  table_data <- rbind(table_data, data.frame(
    Characteristic = "Employment status",
    Total = "--",
    NonProblem = "--",
    Problem = "--",
    t = "N/A",
    p = employment_test$p_value,
    stringsAsFactors = FALSE
  ))
  
  # Get all unique employment statuses and add them as subrows
  employment_statuses <- sort(unique(df$employment))
  for(emp in employment_statuses) {
    if(!is.na(emp) && emp != "") {
      is_emp <- df$employment == emp
      table_data <- rbind(table_data, data.frame(
        Characteristic = paste("  ", emp),
        Total = format_count_pct(is_emp, total_n),
        NonProblem = format_count_pct(is_emp & df[[group_var]] != group_value, non_problem_n),
        Problem = format_count_pct(is_emp & df[[group_var]] == group_value, problem_n),
        t = "-",
        p = "-",
        stringsAsFactors = FALSE
      ))
    }
  }
  
  # Add education
  education_test <- perform_chisq(df$education, df[[group_var]])
  table_data <- rbind(table_data, data.frame(
    Characteristic = "Highest educational level attained",
    Total = "--",
    NonProblem = "--",
    Problem = "--",
    t = "N/A",
    p = education_test$p_value,
    stringsAsFactors = FALSE
  ))
  
  # Get all unique education levels and add them as subrows
  education_levels <- sort(unique(df$education))
  for(edu in education_levels) {
    if(!is.na(edu) && edu != "") {
      is_edu <- df$education == edu
      table_data <- rbind(table_data, data.frame(
        Characteristic = paste("  ", edu),
        Total = format_count_pct(is_edu, total_n),
        NonProblem = format_count_pct(is_edu & df[[group_var]] != group_value, non_problem_n),
        Problem = format_count_pct(is_edu & df[[group_var]] == group_value, problem_n),
        t = "-",
        p = "-",
        stringsAsFactors = FALSE
      ))
    }
  }
  
  # Add gaming time measures header
  table_data <- rbind(table_data, data.frame(
    Characteristic = "Gaming time measures",
    Total = "--",
    NonProblem = "--",
    Problem = "--",
    t = "-",
    p = "-",
    stringsAsFactors = FALSE
  ))
  
  # Add gaming time metrics
  gaming_metrics <- list(
    list(name = "Red box hours; mean (SD)", var = "red_box"),
    list(name = "Green box hours; mean (SD)", var = "green_box"),
    list(name = "Total green & red hours; mean (SD)", var = "green_plus_red"),
    list(name = "Proportion red box hours (%), mean (SD)", var = "red_proportion"),
    list(name = "Typical weekly hours; mean (SD)", var = "typical_weekly_hours")
  )
  
  for(metric in gaming_metrics) {
    t_test <- perform_ttest(df[[metric$var]], df[[group_var]])
    table_data <- rbind(table_data, data.frame(
      Characteristic = metric$name,
      Total = format_mean_sd(df[[metric$var]]),
      NonProblem = format_mean_sd(df[[metric$var]][df[[group_var]] != group_value]),
      Problem = format_mean_sd(df[[metric$var]][df[[group_var]] == group_value]),
      t = t_test$t_value,
      p = t_test$p_value,
      stringsAsFactors = FALSE
    ))
  }
  
  # Add impulsivity header
  table_data <- rbind(table_data, data.frame(
    Characteristic = "Impulsivity (BIS-15)",
    Total = "--",
    NonProblem = "--",
    Problem = "--",
    t = "-",
    p = "-",
    stringsAsFactors = FALSE
  ))
  
  # Add impulsivity metrics
  impulsivity_metrics <- list(
    list(name = "Non-planning; mean (SD)", var = "non_planning"),
    list(name = "Motor; mean (SD)", var = "motor"),
    list(name = "Attentional; mean (SD)", var = "attentional"),
    list(name = "Total score; mean (SD)", var = "total_impulsivity")
  )
  
  for(metric in impulsivity_metrics) {
    t_test <- perform_ttest(df[[metric$var]], df[[group_var]])
    table_data <- rbind(table_data, data.frame(
      Characteristic = metric$name,
      Total = format_mean_sd(df[[metric$var]]),
      NonProblem = format_mean_sd(df[[metric$var]][df[[group_var]] != group_value]),
      Problem = format_mean_sd(df[[metric$var]][df[[group_var]] == group_value]),
      t = t_test$t_value,
      p = t_test$p_value,
      stringsAsFactors = FALSE
    ))
  }
  
  # Add psychological distress header
  table_data <- rbind(table_data, data.frame(
    Characteristic = "Psychological distress (DASS-21)",
    Total = "--",
    NonProblem = "--",
    Problem = "--",
    t = "-",
    p = "-",
    stringsAsFactors = FALSE
  ))
  
  # Add psychological distress metrics
  distress_metrics <- list(
    list(name = "Stress; mean (SD)", var = "stress"),
    list(name = "Anxiety; mean (SD)", var = "anxiety"),
    list(name = "Depression; mean (SD)", var = "depression"),
    list(name = "Total score; mean (SD)", var = "DASS_total")
  )
  
  for(metric in distress_metrics) {
    t_test <- perform_ttest(df[[metric$var]], df[[group_var]])
    table_data <- rbind(table_data, data.frame(
      Characteristic = metric$name,
      Total = format_mean_sd(df[[metric$var]]),
      NonProblem = format_mean_sd(df[[metric$var]][df[[group_var]] != group_value]),
      Problem = format_mean_sd(df[[metric$var]][df[[group_var]] == group_value]),
      t = t_test$t_value,
      p = t_test$p_value,
      stringsAsFactors = FALSE
    ))
  }
  
  # Rename columns to match the desired format
  colnames(table_data) <- c(
    "Characteristics", 
    paste0("Total (N=", total_n, ")"),
    paste0("Non-problem (N=", non_problem_n, ")"),
    paste0(group_value, " (N=", problem_n, ")"),
    "t",
    "p"
  )
  
  return(table_data)
}

# -- Helper Functions --
# Helper function for significance markers
get_sig <- function(p) {
  if (is.na(p)) return("")
  if (p < 0.001) return("***")
  else if (p < 0.01) return("**")
  else if (p < 0.05) return("*")
  else return("")
}

# -- Helper Function for Contingency Table --
create_contingency_table <- function(df, predictor, ref, threshold, disorder_name, predictor_name) {
  pred_binary <- ifelse(predictor > threshold, 1, 0)
  indices <- calc_all_indices(pred_binary, ref)
  
  # Extract TP, TN, FP, FN from the confusion matrix
  TP <- indices$ConfusionMatrix[1, 1] # True Positive
  FN <- indices$ConfusionMatrix[2, 1] # False Negative
  FP <- indices$ConfusionMatrix[1, 2] # False Positive
  TN <- indices$ConfusionMatrix[2, 2] # True Negative
  
  # Calculate totals
  total_positive <- TP + FN
  total_negative <- FP + TN
  total_pred_negative <- TN + FN
  total_pred_positive <- TP + FP
  grand_total <- TP + TN + FP + FN
  
  # Calculate Diagnostic Odds Ratio (DOR) using (TP * TN) / (FP * FN)
  dor <- if (FP == 0 || FN == 0) NA else (TP * TN) / (FP * FN)
  
  # Create the contingency table
  table_data <- data.frame(
    Row = c(
      paste(disorder_name),  # e.g., "GD"
      paste("Non-", disorder_name, sep = ""),  # e.g., "Non-GD"
      "TOTAL",
      "diagnostic accuracy statistics"
    ),
    Pred_Negative = c(
      TN,  # TN: Predicted negative, actual negative
      FN,  # FN: Predicted negative, actual positive
      total_pred_negative,
      paste("NPV =", sprintf("%.2f", indices$NPV))
    ),
    Pred_Positive = c(
      FP,  # FP: Predicted positive, actual negative
      TP,  # TP: Predicted positive, actual positive
      total_pred_positive,
      paste("PPV =", sprintf("%.2f", indices$PPV))
    ),
    Total = c(
      total_negative,
      total_positive,
      grand_total,
      paste("Diagnostic OR =", sprintf("%.2f", dor))
    ),
    Diagnostic_Accuracy = c(
      "",  # Placeholder for first row
      "",  # Placeholder for second row
      "",  # Placeholder for third row
      paste("Se =", sprintf("%.2f", indices$Sensitivity), "\n",
            "Sp =", sprintf("%.2f", indices$Specificity), "\n",
            "LR+ =", sprintf("%.2f", indices$LR_Pos), "\n",
            "LR- =", sprintf("%.2f", indices$LR_Neg))
    ),
    stringsAsFactors = FALSE
  )
  
  # Rename columns to match the desired format
  colnames(table_data) <- c(
    paste(disorder_name, "status"),  # e.g., "GD status"
    paste(predictor_name, "\n<", sprintf("%.1f", threshold), "hours"),  # e.g., "< 9.5 hours"
    paste(predictor_name, "\n>", sprintf("%.1f", threshold), "hours"),  # e.g., "> 9.5 hours"
    "TOTAL",
    "diagnostic accuracy"
  )
  
  return(list(
    table = table_data,
    predictor_name = predictor_name,
    disorder_name = disorder_name
  ))
}

# -- Correlation Analysis --
# Partial Correlation Matrix Function
partial_cor_matrix <- function(df_complete) {
  vars <- c("typical_weekly_hours", "red_box", "green_box", "total_symptoms",
            "GD_status_binary", "IGD_status_binary")
  control_vars <- c("non_planning", "attentional", "motor", "stress", 
                    "anxiety", "depression", "nationality_recoded", 
                    "education_recoded", "employment_recoded")
  
  cor_matrix <- matrix(NA, nrow = length(vars), ncol = length(vars), 
                       dimnames = list(vars, vars))
  p_matrix <- matrix(NA, nrow = length(vars), ncol = length(vars), 
                     dimnames = list(vars, vars))
  
  diag(cor_matrix) <- 1
  
  for (i in 1:(length(vars) - 1)) {
    for (j in (i + 1):length(vars)) {
      tryCatch({
        test <- ppcor::pcor.test(df_complete[[vars[i]]], df_complete[[vars[j]]], 
                                 df_complete[, control_vars])
        cor_val <- test$estimate
        p_val <- test$p.value
        
        cor_matrix[vars[i], vars[j]] <- cor_val
        cor_matrix[vars[j], vars[i]] <- cor_val
        p_matrix[vars[i], vars[j]] <- p_val
        p_matrix[vars[j], vars[i]] <- p_val
      }, error = function(e) {
        warning(paste("Error calculating partial correlation between", 
                      vars[i], "and", vars[j], ":", e$message))
      })
    }
  }
  
  char_matrix <- matrix("", nrow = length(vars), ncol = length(vars), 
                        dimnames = list(vars, vars))
  for (i in 1:length(vars)) {
    for (j in 1:length(vars)) {
      if (i == j) {
        char_matrix[i, j] <- "1.00"
      } else {
        cor_val <- sprintf("%.2f", cor_matrix[i, j])
        sig <- get_sig(p_matrix[i, j])
        char_matrix[i, j] <- paste0(cor_val, sig)
      }
    }
  }
  
  return(char_matrix)
}

# Zero-Order Correlation Matrix Function
zero_order_cor_matrix <- function(df_complete) {
  vars <- c("typical_weekly_hours", "red_box", "green_box", "total_symptoms",
            "GD_status_binary", "IGD_status_binary", "non_planning",
            "attentional", "motor", "stress", "anxiety", "depression", 
            "age", "nationality_recoded", "employment_recoded", "education_recoded")
  
  cor_matrix <- cor(df_complete[, vars], use = "pairwise.complete.obs")
  
  p_matrix <- matrix(NA, nrow = length(vars), ncol = length(vars), 
                     dimnames = list(vars, vars))
  
  for (i in 1:(length(vars) - 1)) {
    for (j in (i + 1):length(vars)) {
      tryCatch({
        test <- cor.test(df_complete[[vars[i]]], df_complete[[vars[j]]])
        p_val <- test$p.value
        p_matrix[vars[i], vars[j]] <- p_val
        p_matrix[vars[j], vars[i]] <- p_val
      }, error = function(e) {
        warning(paste("Error calculating correlation between", 
                      vars[i], "and", vars[j], ":", e$message))
      })
    }
  }
  
  char_matrix <- matrix("", nrow = length(vars), ncol = length(vars), 
                        dimnames = list(vars, vars))
  for (i in 1:length(vars)) {
    for (j in 1:length(vars)) {
      if (i == j) {
        char_matrix[i, j] <- "1.00"
      } else {
        cor_val <- sprintf("%.2f", cor_matrix[i, j])
        sig <- get_sig(p_matrix[i, j])
        char_matrix[i, j] <- paste0(cor_val, sig)
      }
    }
  }
  
  return(char_matrix)
}

# -- ROC Analysis --
# Compute diagnostic accuracy indices
calc_all_indices <- function(pred, ref) {
  TP <- sum(pred == 1 & ref == 1, na.rm = TRUE)
  TN <- sum(pred == 0 & ref == 0, na.rm = TRUE)
  FP <- sum(pred == 1 & ref == 0, na.rm = TRUE)
  FN <- sum(pred == 0 & ref == 1, na.rm = TRUE)
  
  cm <- matrix(c(TP, FP, FN, TN), nrow = 2, byrow = TRUE,
               dimnames = list(Predicted = c("Positive", "Negative"),
                               Actual = c("Positive", "Negative")))
  
  sensitivity <- round(TP / (TP + FN), 2)
  specificity <- round(TN / (TN + FP), 2)
  ppv <- round(TP / (TP + FP), 2)
  npv <- round(TN / (TN + FN), 2)
  
  lr_pos <- round(sensitivity / (1 - specificity), 2)
  lr_neg <- round((1 - sensitivity) / specificity, 2)
  
  cui_pos <- round(sensitivity * ppv, 2)
  cui_neg <- round(specificity * npv, 2)
  
  if (is.nan(ppv)) ppv <- 0
  if (is.nan(npv)) npv <- 0
  if (is.infinite(lr_pos) || is.nan(lr_pos)) lr_pos <- NA
  if (is.infinite(lr_neg) || is.nan(lr_neg)) lr_neg <- NA
  
  return(list(
    ConfusionMatrix = cm,
    Sensitivity = sensitivity,
    Specificity = specificity,
    PPV = ppv,
    NPV = npv,
    LR_Pos = lr_pos,
    LR_Neg = lr_neg,
    CUI_Pos = cui_pos,
    CUI_Neg = cui_neg
  ))
}

calc_indices_with_threshold <- function(predictor, ref, threshold, predictor_name, threshold_type, disorder_type, predictor_order) {
  pred_binary <- ifelse(predictor > threshold, 1, 0)
  indices <- calc_all_indices(pred_binary, ref)
  
  # Extract TP, TN, FP, FN from the confusion matrix
  TP <- indices$ConfusionMatrix[1, 1] # True Positive
  FN <- indices$ConfusionMatrix[2, 1] # False Negative
  FP <- indices$ConfusionMatrix[1, 2] # False Positive
  TN <- indices$ConfusionMatrix[2, 2] # True Negative
  
  # Calculate Diagnostic Odds Ratio (DOR) using (TP * TN) / (FP * FN)
  dor <- if (FP == 0 || FN == 0) NA else round((TP * TN) / (FP * FN), 2)
  
  return(list(
    Predictor = predictor_name, 
    Disorder = disorder_type,
    Cut_Off = round(threshold, 2),
    ConfusionMatrix = indices$ConfusionMatrix,
    Sensitivity = indices$Sensitivity,
    Specificity = indices$Specificity,
    PPV = indices$PPV,
    NPV = indices$NPV,
    LR_Pos = indices$LR_Pos,
    LR_Neg = indices$LR_Neg,
    CUI_Pos = indices$CUI_Pos,
    CUI_Neg = indices$CUI_Neg,
    AUC = auc(roc(ref, predictor, quiet = TRUE)),
    Diagnostic_OR = dor,
    Predictor_Order = predictor_order
  ))
}

# Function to plot ROC curves in a 2x2 grid with Youden's Index
plot_roc_curves <- function(data, predictors, outcome, file_path, disorder_name) {
  png(file_path, width = 2126, height = 2126, res = 300)  # 180 mm at 300 DPI
  par(mfrow = c(2, 2), mar = c(4, 4, 2, 1))
  
  colors <- c("red", "darkgreen", "blue", "purple")
  predictor_names <- c("Red Box", "Green Box", "% Red Box Hours", "Typical Weekly Hours")
  
  for (i in seq_along(predictors)) {
    roc_obj <- roc(data[[outcome]], data[[predictors[i]]], quiet = TRUE)
    auc_val <- round(auc(roc_obj), 3)
    
    # Get Youden's Index coordinates
    youden_coords <- coords(roc_obj, "best", ret = c("threshold", "specificity", "sensitivity"), best.method = "youden")
    specificity <- youden_coords$specificity
    sensitivity <- youden_coords$sensitivity
    
    # Plot ROC curve
    plot(roc_obj, col = colors[i], lwd = 2, 
         main = paste(predictor_names[i], "(AUC =", auc_val, ")"),
         xlab = "1 - Specificity", ylab = "Sensitivity")
  }
  
  par(mfrow = c(1, 1))
  dev.off()
}

# -- Main Analysis Function --
run_analysis <- function(user_dir = NULL, data_file = "red_box.xlsx") {
  # Set up project directories
  project_dirs <- setup_project(data_file = data_file, user_dir = user_dir)
  
  tryCatch({
    # Check if data file exists and is readable
    if (file.exists(project_dirs$data_path)) {
      message("Reading data file from: ", project_dirs$data_path)
      # Try to read Excel file
      tryCatch({
        df <- read_excel(project_dirs$data_path)
      }, error = function(e) {
        stop(paste("Error reading Excel file:", e$message, 
                   "\nEnsure the file exists and is a valid Excel file."))
      })
    } else {
      stop(paste("Data file not found:", project_dirs$data_path))
    }
    
    # Process data
    message("Processing data...")
    df <- process_data(df)
    
    # Create a complete cases dataset for correlation analysis
    df_complete <- df[complete.cases(df[, c(
      "typical_weekly_hours", "red_box", "green_box", "IGD_status_binary", 
      "GD_status_binary", "total_symptoms", "non_planning", "attentional", 
      "motor", "total_impulsivity", "stress", "anxiety", "depression", 
      "DASS_total", "age", "nationality_recoded", "employment_recoded", 
      "education_recoded"
    )]), ]
    
    # Generate descriptive statistics table for GD
    message("Generating descriptive statistics tables...")
    table_1 <- create_descriptive_table(df, group_var = "GD_status", group_value = "GD", table_title = "Sample Characteristics by GD Status")
    
    # Generate supplementary descriptive statistics table for IGD
    supp_table_1 <- create_descriptive_table(df, group_var = "IGD_status", group_value = "IGD", table_title = "Sample Characteristics by IGD Status")
    
    # Save descriptive statistics tables
    write.csv(table_1, 
              file.path(project_dirs$dir_paths$manuscript_files, "Table 1 - Sample Characteristics.csv"), 
              row.names = FALSE)
    write.csv(supp_table_1, 
              file.path(project_dirs$dir_paths$supplementary_files, "Supplementary Table 1 - Sample Characteristics by IGD Status.csv"), 
              row.names = FALSE)
    
    # Compute correlation matrices
    message("Computing correlation matrices...")
    zero_cor_mat <- zero_order_cor_matrix(df_complete)
    partial_cor_mat <- partial_cor_matrix(df_complete)
    
    # Save correlation results
    write.csv(zero_cor_mat, file.path(project_dirs$dir_paths$supplementary_files, 
                                      "Supplementary Table 2 - Zero-order correlation matrix.csv"), 
              row.names = TRUE)
    write.csv(partial_cor_mat, file.path(project_dirs$dir_paths$manuscript_files, 
                                         "Table 2 - Partial-order correlation matrix.csv"), 
              row.names = TRUE)
    
    # Define predictors with their order
    predictors <- list(
      list(name = "Red Box (total hrs)", data = df$red_box, order = 1),
      list(name = "Green Box (total hrs)", data = df$green_box, order = 2),
      list(name = "Green + red hours", data = df$green_plus_red, order = 3),
      list(name = "Proportion red box", data = df$red_proportion, order = 4),
      list(name = "Typical weekly hours", data = df$typical_weekly_hours, order = 5)
    )
    
    # Define disorder types
    disorders <- list(
      list(name = "GD", data = df$GD_status_binary),
      list(name = "IGD", data = df$IGD_status_binary)
    )
    
    # Calculate thresholds and indices
    message("Performing ROC analysis...")
    diag_results <- list()
    for (d in disorders) {
      for (p in predictors) {
        if (sum(d$data, na.rm = TRUE) >= 2) {
          roc_obj <- roc(d$data, p$data, quiet = TRUE)
          auc_value <- round(auc(roc_obj), 2)
          roc_threshold <- coords(roc_obj, "best", ret = "threshold", best.method = "youden")$threshold
          
          result_key <- paste(d$name, p$name, "ROC", sep = "_")
          diag_results[[result_key]] <- calc_indices_with_threshold(
            p$data, d$data, roc_threshold, p$name, "ROC-Optimized (Youden)", d$name, p$order
          )
          
          diag_results[[result_key]]$AUC <- auc_value
        } else {
          warning(paste("Not enough positive cases for", d$name, "to perform ROC analysis with", p$name))
        }
      }
    }
    
    # Create diagnostic accuracy table
    table_2 <- do.call(rbind, lapply(diag_results, function(x) {
      data.frame(
        Disorder = x$Disorder,
        Predictor = x$Predictor,
        Cut_Off = x$Cut_Off,
        Sensitivity = x$Sensitivity,
        Specificity = x$Specificity,
        PPV = x$PPV,
        NPV = x$NPV,
        LR_Pos = x$LR_Pos,
        LR_Neg = x$LR_Neg,
        CUI_Pos = x$CUI_Pos,
        CUI_Neg = x$CUI_Neg,
        AUC = x$AUC,
        Diagnostic_OR = x$Diagnostic_OR,
        Predictor_Order = x$Predictor_Order,
        stringsAsFactors = FALSE
      )
    }))
    
    # Order table by disorder and predictor
    table_2 <- table_2[order(table_2$Disorder, table_2$Predictor_Order), ]
    
    # Remove Disorder and Predictor_Order columns
    table_2 <- subset(table_2, select = -c(Predictor_Order))
    
    # Create header rows
    header_gd <- data.frame(
      Disorder = "GD",
      Predictor = "ICD-11 gaming disorder",
      Cut_Off = "",
      Sensitivity = "",
      Specificity = "",
      PPV = "",
      NPV = "",
      LR_Pos = "",
      LR_Neg = "",
      CUI_Pos = "",
      CUI_Neg = "",
      AUC = "",
      Diagnostic_OR = "",
      stringsAsFactors = FALSE
    )
    
    header_igd <- data.frame(
      Disorder = "IGD",
      Predictor = "DSM-5 internet gaming disorder",
      Cut_Off = "",
      Sensitivity = "",
      Specificity = "",
      PPV = "",
      NPV = "",
      LR_Pos = "",
      LR_Neg = "",
      CUI_Pos = "",
      CUI_Neg = "",
      AUC = "",
      Diagnostic_OR = "",
      stringsAsFactors = FALSE
    )
    
    # Split table into GD and IGD sections
    gd_rows <- table_2[table_2$Disorder == "GD", ]
    igd_rows <- table_2[table_2$Disorder == "IGD", ]
    
    # Combine with headers
    final_table <- rbind(
      header_gd,
      gd_rows,
      header_igd,
      igd_rows
    )
    
    # Save diagnostic accuracy table
    write.csv(final_table, file.path(project_dirs$dir_paths$manuscript_files, 
                                     "Table 3 - Diagnostic accuracy indices.csv"), 
              row.names = FALSE)
    
    # Add contingency tables for GD and IGD
    message("Generating contingency tables...")
    contingency_tables <- list()
    
    for (d in disorders) {
      for (p in predictors) {
        if (sum(d$data, na.rm = TRUE) >= 2) {
          roc_obj <- roc(d$data, p$data, quiet = TRUE)
          roc_threshold <- coords(roc_obj, "best", ret = "threshold", best.method = "youden")$threshold
          
          contingency_table <- create_contingency_table(
            df, p$data, d$data, roc_threshold, d$name, p$name
          )
          
          result_key <- paste(d$name, p$name, "Contingency", sep = "_")
          contingency_tables[[result_key]] <- contingency_table
        }
      }
    }
    
    # Save contingency tables
    for (ct in contingency_tables) {
      file_name <- sprintf("Supplementary Table - Contingency %s - %s.csv", ct$disorder_name, ct$predictor_name)
      write.csv(ct$table, 
                file.path(project_dirs$dir_paths$supplementary_files, file_name), 
                row.names = FALSE)
    }
    
    # Plot ROC curves
    roc_predictors <- c("red_box", "green_box", "red_proportion", "typical_weekly_hours")
    plot_roc_curves(df_complete, roc_predictors, "GD_status_binary", 
                    file.path(project_dirs$dir_paths$manuscript_files, "Figure 1 - ROC Curves for ICD-11.png"),
                    "ICD-11 gaming disorder")
    plot_roc_curves(df_complete, roc_predictors, "IGD_status_binary", 
                    file.path(project_dirs$dir_paths$supplementary_files, "Supplementary Figure 2 - ROC Curves for DSM-5.png"),
                    "DSM-5 internet gaming disorder")
    
    # Print completion message with output locations
    message("\n=== ANALYSIS COMPLETE ===")
    message("Output files have been saved to the following locations:")
    message("- Manuscript Tables and Figures: ", project_dirs$dir_paths$manuscript_files)
    message("- Supplementary Files: ", project_dirs$dir_paths$supplementary_files)
    
    # Return results
    return(list(
      descriptive_stats = table_1,
      supplementary_stats = supp_table_1,
      zero_order = zero_cor_mat,
      partial_order = partial_cor_mat,
      diagnostic_accuracy = final_table,
      contingency_tables = contingency_tables,
      output_locations = list(
        tables = project_dirs$dir_paths$manuscript_files,
        supplementary = project_dirs$dir_paths$supplementary_files
      )
    ))
  }, error = function(e) {
    stop(paste("Error in analysis:", e$message))
  })
}

# Execute analysis
results <- run_analysis()
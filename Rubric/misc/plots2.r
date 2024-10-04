library(rvg)
library(dplyr)
library(tidyr)
library(openxlsx)
library(ggthemes)
library(ggplot2)
library(officer)


file.path <- "H:/BB5G/Main_table/"
setwd(file.path)
df <- read.csv(paste0(file.path, "CIP Brian for table .csv"), header = TRUE)%>%
  mutate(CIP = as.character(CIP))
# 
# head(df)
# names(df)
# str(df)

cips <- read.csv(paste0(file.path, "master_CIP.csv"),header = TRUE)%>%
  mutate(CIP = as.character(CIP))
head(cips)


df <- df%>%
  left_join(cips, by = 'CIP')%>%
  mutate(CIP = paste0(CIPDescription, " (", CIP, ")"))%>%
  dplyr::select(-CIPDescription)%>%
  mutate(JobsOhio.Region = recode(JobsOhio.Region,
                                  "NW" = "Regional Growth Partnership",
                                  "SW" = "REDI Cincinnati",
                                  "NE" = "Team NEO",
                                  "C" = "OneColumbus",
                                  "W" = "Dayton Development Coalition",
                                  "SE" = "Appalachian Partnership"))

# Function to print DataFrame to Excel
print_df_to_excel <- function(data, region) {
  # Create Excel file
  wb <- createWorkbook()
  addWorksheet(wb, region)
  
  # Write DataFrame to Excel
  writeData(wb, region, data)
  
  # Save Excel file
  saveWorkbook(wb, paste0(region, ".xlsx"), overwrite = TRUE)
}

# Filter unique regions
unique_regions <- unique(df$JobsOhio.Region)

# Iterate over unique regions and print DataFrame to Excel
for (region in unique_regions) {
  region_data <- subset(df, JobsOhio.Region == region)
  
  # Call function to print DataFrame to Excel
  print_df_to_excel(region_data, region)
}


# 
# ###OLD
# # Define a custom color palette with enough colors
# palette <- c("#E41A1C", "#377EB8", "#4DAF4A", "#984EA3", "#FF7F00", "#FFFF33", "#A65628", "#F781BF", "#999999", "#66C2A5", 
#             "#FC8D62", "#8DA0CB", "#E78AC3", "#A6D854", "#FFD92F", "#E5C494", "#B3B3CC", "#8DD3C7", "#FFFFB3", "#BEBADA", 
#             "#FB8072", "#80B1D3", "#FDB462", "#B3DE69", "#FCCDE5", "#D9D9D9", "#BC80BD", "#CCEBC5", "#FFED6F")
# 
# # Function to plot line graphs
# plot_line_graphs <- function(data, region, print_legend = FALSE) {
#   # Convert Year column to numeric
#   data$Year <- as.numeric(data$Year)
#   
#   # Create a line plot with points
#   p <- ggplot(data, aes(x = Year, y = value, group = CIP, color = CIP)) +
#     geom_line() +
#     geom_point() +
#     geom_text(data = subset(data, Year == max(Year)), aes(label = CIP),
#               nudge_x = 0.1, hjust = 0) +
#     labs(x = "Year", y = "Graduates", title = region) +
#     theme_bw() +
#     theme(legend.position = "none",
#           panel.border = element_blank(),
#     panel.grid.major.x = element_blank(),
#   panel.grid.minor.x = element_blank()) +
#     scale_color_manual(values = palette)  # Use custom color palette
#   
#   # Convert the ggplot object to a rvg object
#   p_dml <- dml(ggobj = p)
#   
#   # Create a PowerPoint slide
#   slide <- read_pptx() %>%
#     add_slide() %>%
#     ph_with(p_dml, ph_location())
#   
#   # Save the PowerPoint file
#   if (region == "Ohio") {
#     # Print the legend to a separate PowerPoint
#     legend_slide <- read_pptx() %>%
#       add_slide() %>%
#       ph_with(p_dml + theme(legend.position = "bottom"), ph_location())
#     print(legend_slide, target = "Ohio_legend.pptx")
#   } else {
#     print(slide, target = paste0("OhioJobs_", region, "_line_plot.pptx"))
#   }
# }
# # Iterate over unique regions and plot line graphs
# for (region in unique_regions) {
#   region_data <- subset(df, JobsOhio.Region == region)
#   
#   # Reshape the data to long format for plotting
#   df_long <- region_data %>%
#     select(CIP, X2019, X2020, X2021) %>%
#     rename(`2019` = X2019, `2020` = X2020, `2021` = X2021) %>%
#     mutate(CIP = gsub(".*\\((\\d+\\.\\d+)\\).*", "\\1", CIP)) %>%
#     pivot_longer(cols = c(`2019`, `2020`, `2021`), names_to = "Year", values_to = "value") %>%
#     mutate(Year = as.numeric(Year))  # Convert Year column to numeric
#   
#   if (region == "Ohio") {
#     # Call function to plot line graphs and print the legend to a separate PowerPoint
#     plot_line_graphs(df_long, paste0("OhioJobs_", region, "_Region"), print_legend = TRUE)
#   } else {
#     # Call function to plot line graphs without printing the legend
#     plot_line_graphs(df_long, paste0("OhioJobs_", region, "_Region"))
#   }
# }
# 
# 
# ##############################
# #Version that prints two line charts, one over 500 and one below for each area
# #########################################
# 
# ###OLD II
# # Define a custom color palette with more visible colors on white paper
# palette <- c("#1F77B4", "#FF7F0E", "#2CA02C", "#D62728", "#9467BD", "#8C564B", "#8B5A00", "#E377C2", "#7F7F7F", "#B4EEB4", "#17BECF",
#                       "#FFD700", "#FF1493", "#00BFFF", "#FF4500", "#8B008B", "#32CD32", "#FF8C00")
#                       
# # Function to plot line graphs with separate chart for CIPs over 500
# plot_line_graphs <- function(data, region) {
#   # Convert Year column to numeric
#   data$Year <- as.numeric(data$Year)
#   
#   # Separate data for CIPs over 500 and under 500
#   data_over_500 <- subset(data, value >= 500)
#   data_under_500 <- subset(data, value <= 500)
#   
#   # Create line plots for CIPs over 500 (stacked on top)
#   p_over_500 <- ggplot(data_over_500, aes(x = Year, y = value, group = CIP, color = CIP)) +
#     geom_line() +
#     geom_point() +
#     geom_text(data = subset(data_over_500, Year == max(Year)), aes(label = CIP),
#               nudge_x = 0.1, hjust = 0, size = 3) +  # Set CIP text size to 9pt
#     geom_text(data = subset(data_over_500, value == max(value)), aes(label = value),
#               nudge_y = 500, hjust = 0, color = "black") +  # Add black text for highest value
#     labs(x = NULL, y = "Graduates", title = paste(region, "CIPs", "(over 500 graduates)")) +
#     theme_bw() +
#     theme(
#       legend.position = "none",
#       panel.border = element_rect(color = "black", fill = NA, size = 0.5),  # Add left and right borders
#       panel.grid.major.y = element_line(color = "gray70"),  # Set y-axis gridlines color
#       panel.grid.minor.x = element_blank(),  # Remove x-axis gridlines
#       axis.text.x = element_blank(),
#       axis.ticks.x = element_blank(),
#       axis.text.y = element_text(size = 9)  # Set y-axis text size to 9pt
#     ) +
#     scale_color_manual(values = palette) +
#     scale_y_continuous(labels = scales::comma, breaks = seq(0, ceiling(max(data_over_500$value) / 500) * 500, by = 500))  # Adjust y-axis breaks
#   
#   # Create line plots for CIPs under 500
#   p_under_500 <- ggplot(data_under_500, aes(x = Year, y = value, group = CIP, color = CIP)) +
#     geom_line() +
#     geom_point() +
#     geom_text(data = subset(data_under_500, Year == max(Year)), aes(label = CIP),
#               nudge_x = 0.1, hjust = 0, size = 3) +  # Set CIP text size to 9pt
#     labs(x = "Year", y = "Graduates", title = "(under 500 graduates)") +
#     theme_bw() +
#     theme(
#       legend.position = "none",
#       panel.border = element_rect(color = "black", fill = NA, size = 0.5),  # Add left and right borders
#       panel.grid.major.y = element_line(color = "gray70"),  # Set y-axis gridlines color
#       panel.grid.minor.x = element_blank(),  # Remove x-axis gridlines
#       axis.text.y = element_text()
#     ) +
#     scale_color_manual(values = palette) +
#     scale_x_continuous(breaks = c(2019, 2020, 2021), labels = c(2019, 2020, 2021)) +
#     scale_y_continuous(labels = scales::comma, breaks = seq(0, 600, by = 100)) +
#     labs(caption = "Values represent the number of graduates")
#   
#   # Convert the ggplot objects to rvg objects
#   p_dml_over_500 <- dml(ggobj = p_over_500)
#   p_dml_under_500 <- dml(ggobj = p_under_500)
#   
#   
#   
#   # Create a PowerPoint slide
#   slide <- read_pptx() %>%
#     add_slide() %>%
#     ph_with(p_dml_over_500, ph_location(left = 0, top = 0, width = 9.5, height = 3.5)) %>%
#     ph_with(p_dml_under_500, ph_location(left = 0, top = 7.5, width = 9.5, height = 3.5))
#   
#   # Save the PowerPoint file
#   print(slide, target = paste0("OhioJobs_", region, "_line_plot.pptx"))
# }
# 
# # Iterate over unique regions and plot line graphs
# for (region in unique_regions) {
#   region_data <- subset(df, JobsOhio.Region == region)
#   # Reshape the data to long format for plotting
#   df_long <- region_data %>%
#     select(CIP, X2019, X2020, X2021) %>%
#     rename(`2019` = X2019, `2020` = X2020, `2021` = X2021) %>%
#     mutate(CIP = gsub(".*\\((\\d+\\.\\d+)\\).*", "\\1", CIP)) %>%
#     pivot_longer(cols = c(`2019`, `2020`, `2021`), names_to = "Year", values_to = "value") %>%
#     mutate(Year = as.numeric(Year))  # Convert Year column to numeric
#   
#   # Call function to plot line graphs
#   plot_line_graphs(df_long, paste0("OhioJobs_", region, "_Region"))
# }
# 



##Version II Actual
# Function to plot line graphs with separate chart for CIPs over 500
plot_line_graphs <- function(data, region) {
  # Convert Year column to numeric
  data$Year <- as.numeric(data$Year)
  
  # Separate data for CIPs over 500 and under 500
  data_over_500 <- subset(data, value >= 500)
  data_under_500 <- subset(data, value <= 500)
  
  # Get unique CIPs from the total dataframe
  unique_cips <- unique(data$CIP)
  
  # Load the paletteer package
  library(paletteer)
  
  # Define the palettepaletteer_c("grDevices::rainbow", 30)
  custom_palette <- paletteer_c("grDevices::Dark 3", 30)
  
  # Assign colors to unique CIPs using the custom palette
  cip_colors <- custom_palette[1:length(unique_cips)]
  
  # Create a color mapping for CIPs
  color_mapping <- setNames(cip_colors, unique_cips)
  
  
  # Create line plots for CIPs over 500 (stacked on top)
  p_over_500 <- ggplot(data_over_500, aes(x = Year, y = value, group = CIP, color = CIP)) +
    geom_line() +
    geom_point() +
    geom_text(data = subset(data_over_500, Year == max(Year)), aes(label = CIP),
              nudge_x = 0.1, hjust = 0, size = 3) +  # Set CIP text size to 9pt
    geom_text(data = subset(data_over_500, value == max(value)), aes(label = value),
              nudge_y = 500, hjust = 0, color = "black") + 
    geom_smooth(method = "lowess", se = FALSE, color = "red")+# Add black text for highest value
    labs(x = NULL, y = "Graduates", title = paste(region, "CIPs", "(Over 500 Graduates)")) +
    theme_bw() +
    theme(
      legend.position = "none",
      panel.border = element_rect(color = "black", fill = NA, size = 0.5),  # Add left and right borders
      panel.grid.major.y = element_line(color = "gray70"),  # Set y-axis gridlines color
      panel.grid.minor.x = element_blank(),  # Remove x-axis gridlines
      axis.text.x = element_blank(),
      axis.ticks.x = element_blank(),
      axis.text.y = element_text(size = 9)  # Set y-axis text size to 9pt
    ) +
    scale_color_manual(values = color_mapping) +
    scale_y_continuous(labels = scales::comma, breaks = seq(0, ceiling(max(data_over_500$value) / 500) * 500, by = 500))  # Adjust y-axis breaks
  
  
  # Create line plots for CIPs under 500
  p_under_500 <- ggplot(data_under_500, aes(x = Year, y = value, group = CIP, color = CIP)) +
    geom_line() +
    geom_point() +
    geom_text(data = subset(data_under_500, Year == max(Year)), aes(label = CIP),
              nudge_x = 0.1, hjust = 0, size = 3) +  # Set CIP text size to 9pt
    geom_smooth(method = "lowess", se = FALSE, color = "red")+
    labs(x = "Year", y = "Graduates", title = "(Under 500 Graduates)") +
    theme_bw() +
    theme(
      legend.position = "none",
      panel.border = element_rect(color = "black", fill = NA, size = 0.5),  # Add left and right borders
      panel.grid.major.y = element_line(color = "gray70"),  # Set y-axis gridlines color
      panel.grid.minor.x = element_blank(),  # Remove x-axis gridlines
      axis.text.y = element_text()
    ) +
    scale_color_manual(values = color_mapping) +
    scale_x_continuous(breaks = c(2019, 2020, 2021), labels = c(2019, 2020, 2021)) +
    scale_y_continuous(labels = scales::comma, breaks = seq(0, 600, by = 100))
  
  # Convert the ggplot objects to rvg objects
  p_dml_over_500 <- dml(ggobj = p_over_500)
  p_dml_under_500 <- dml(ggobj = p_under_500)
  
  
  
  # Create a PowerPoint slide
  slide <- read_pptx() %>%
    add_slide() %>%
    ph_with(p_dml_over_500, ph_location(left = 0, top = 0, width = 9.5, height = 3.5)) %>%
    ph_with(p_dml_under_500, ph_location(left = 0, top = 7.5, width = 9.5, height = 3.5))
  
  # Save the PowerPoint file
  print(slide, target = paste0("OhioJobs_", region, "_line_plot.pptx"))
}

# Iterate over unique regions and plot line graphs
for (region in unique_regions) {
  region_data <- subset(df, JobsOhio.Region == region)
  # Reshape the data to long format for plotting
  df_long <- region_data %>%
    dplyr::select(CIP, X2019, X2020, X2021) %>%
    rename(`2019` = X2019, `2020` = X2020, `2021` = X2021) %>%
    mutate(CIP = gsub(".*\\((\\d+\\.\\d+)\\).*", "\\1", CIP)) %>%
    pivot_longer(cols = c(`2019`, `2020`, `2021`), names_to = "Year", values_to = "value") %>%
    mutate(Year = as.numeric(Year))  # Convert Year column to numeric
  
  # Call function to plot line graphs
  plot_line_graphs(df_long, paste0("OhioJobs_", region, "_Region"))
}

  
  
##############################
#Version that prints one line chart per area with a break on the x-axis. 
#########################################
# 
# 
# # Function to plot line graphs with adjusted y-axis
# plot_line_graphs <- function(data, region) {
#   # Convert Year column to numeric
#   data$Year <- as.numeric(data$Year)
#   
#   # Create line plot with adjusted y-axis
#   p <- ggplot(data, aes(x = Year, y = value, group = CIP, color = CIP)) +
#     geom_line() +
#     geom_point() +
#     geom_text(data = subset(data, Year == max(Year)), aes(label = CIP),
#               nudge_x = 0.1, hjust = 0) +
#     geom_hline(yintercept = 500, linetype = "dashed", color = "red") +
#     labs(x = "Year", y = "Graduates", title = paste(region, "CIPs")) +
#     theme_bw() +
#     theme(legend.position = "none",
#           panel.border = element_blank(),
#           panel.grid.major.x = element_blank(),
#           panel.grid.minor.x = element_blank()) +
#     scale_color_manual(values = palette) +  # Use custom color palette
#     scale_x_continuous(breaks = c(2019, 2020, 2021), labels = c(2019, 2020, 2021), expand = c(0, 0)) +
#     scale_y_continuous(breaks = c(0, 100, 200, 300, 400, 500, max(data$value)),
#                        limits = c(0, max(data$value) * 1.2),
#                        expand = expansion(mult = c(0.75, 0)),
#                        oob = scales::oob_keep) +
#     coord_cartesian(clip = "off") +
#     guides(color = guide_legend(nrow = 2))
#   
#   # Convert the ggplot object to an rvg object
#   p_dml <- dml(ggobj = p)
#   
#   # Create a PowerPoint slide
#   slide <- read_pptx() %>%
#     add_slide() %>%
#     ph_with(p_dml, ph_location(left = 0, top = 0, width = 10, height = 7.5))
#   
#   # Save the PowerPoint file
#   print(slide, target = paste0("OhioJobs_", region, "_line_plot.pptx"))
# }
# 
# # Iterate over unique regions and plot line graphs
# for (region in unique_regions) {
#   region_data <- subset(df, JobsOhio.Region == region)
#   # Reshape the data to long format for plotting
#   df_long <- region_data %>%
#     select(CIP, X2019, X2020, X2021) %>%
#     rename(`2019` = X2019, `2020` = X2020, `2021` = X2021) %>%
#     mutate(CIP = gsub(".*\\((\\d+\\.\\d+)\\).*", "\\1", CIP)) %>%
#     pivot_longer(cols = c(`2019`, `2020`, `2021`), names_to = "Year", values_to = "value") %>%
#     mutate(Year = as.numeric(Year))  # Convert Year column to numeric
#   
#   # Call function to plot line graphs
#   plot_line_graphs(df_long, paste0("OhioJobs_", region, "_Region"))
# }
# 
# #Attempt II at the above ######################################################
# # Function to plot line graphs with combined chart for CIPs over 500 and under 500
# plot_line_graphs <- function(data, region) {
#   # Convert Year column to numeric
#   data$Year <- as.numeric(data$Year)
#   
#   # Create line plot with combined CIPs
#   p <- ggplot(data, aes(x = Year, y = value, group = CIP, color = CIP)) +
#     geom_line() +
#     geom_point() +
#     geom_text(data = subset(data, Year == max(Year)), aes(label = CIP), nudge_x = 0.1, hjust = 0) +
#     labs(x = "Year", y = "Graduates", title = paste(region, "CIPs")) +
#     theme_bw() +
#     theme(legend.position = "none",
#           panel.border = element_blank(),
#           panel.grid.major.x = element_blank(),
#           panel.grid.minor.x = element_blank()) +
#     scale_color_manual(values = palette) +
#     scale_x_continuous(breaks = c(2019, 2020, 2021), labels = c(2019, 2020, 2021)) +
#     geom_hline(yintercept = 500, color = "black", linetype = "dashed") +
#     coord_cartesian(clip = "off") +
#     guides(color = guide_legend(nrow = 2))+
#     facet_zoom(ylim = c(0, 500), zoom.data = ifelse(a <= 10, NA, FALSE))
#   
#   # Convert the ggplot object to an rvg object
#   p_dml <- dml(ggobj = p)
#   
#   # Create a PowerPoint slide
#   slide <- read_pptx() %>%
#     add_slide() %>%
#     ph_with(p_dml, ph_location(left = 0, top = 0, width = 10, height = 7.5))
#   
#   # Save the PowerPoint file
#   print(slide, target = paste0("OhioJobs_", region, "_line_plot.pptx"))
# }
# 
# # Iterate over unique regions and plot line graphs
# for (region in unique_regions) {
#   region_data <- subset(df, JobsOhio.Region == region)
#   # Reshape the data to long format for plotting
#   df_long <- region_data %>%
#     select(CIP, X2019, X2020, X2021) %>%
#     rename(`2019` = X2019, `2020` = X2020, `2021` = X2021) %>%
#     mutate(CIP = gsub(".*\\((\\d+\\.\\d+)\\).*", "\\1", CIP)) %>%
#     pivot_longer(cols = c(`2019`, `2020`, `2021`), names_to = "Year", values_to = "value") %>%
#     mutate(Year = as.numeric(Year))  # Convert Year column to numeric
#   
#   # Call function to plot line graphs
#   plot_line_graphs(df_long, paste0("OhioJobs_", region, "_Region"))
# }
# 
# 
# #ATTEMPT III
# 
# # Function to plot line graphs with combined chart for CIPs
# plot_line_graphs <- function(data, region) {
#   # Convert Year column to numeric
#   data$Year <- as.numeric(data$Year)
#   
#   # Create line plot with combined CIPs
#   p <- ggplot(data, aes(x = Year, y = value, group = CIP, color = CIP)) +
#     geom_line() +
#     geom_point() +
#     geom_text(data = subset(data, Year == max(Year)), aes(label = CIP), nudge_x = 0.1, hjust = 0) +
#     labs(x = "Year", y = "Graduates", title = paste(region, "CIPs")) +
#     theme_bw() +
#     theme(legend.position = "none",
#           panel.border = element_blank(),
#           panel.grid.major.x = element_blank(),
#           panel.grid.minor.x = element_blank()) +
#     scale_color_manual(values = palette) +
#     scale_x_continuous(breaks = c(2019, 2020, 2021), labels = c(2019, 2020, 2021)) +
#     scale_y_continuous(expand = c(0, 0), limits = c(0, max(data$value) * 1.2),
#                        breaks = c(0, 4, 30, max(data$value)),
#                        labels = c(0, 4, 30, max(data$value))) +
#     geom_hline(yintercept = 500, color = "black", linetype = "dashed") +
#     coord_cartesian(clip = "off") +
#     guides(color = guide_legend(nrow = 2))
#   
#   # Convert the ggplot object to an rvg object
#   p_dml <- dml(ggobj = p)
#   
#   # Create a PowerPoint slide
#   slide <- read_pptx() %>%
#     add_slide() %>%
#     ph_with(p_dml, ph_location(left = 0, top = 0, width = 10, height = 7.5))
#   
#   # Save the PowerPoint file
#   print(slide, target = paste0("OhioJobs_", region, "_line_plot.pptx"))
# }
# 
# # Iterate over unique regions and plot line graphs
# for (region in unique_regions) {
#   region_data <- subset(df, JobsOhio.Region == region)
#   # Reshape the data to long format for plotting
#   df_long <- region_data %>%
#     select(CIP, X2019, X2020, X2021) %>%
#     rename(`2019` = X2019, `2020` = X2020, `2021` = X2021) %>%
#     mutate(CIP = gsub(".*\\((\\d+\\.\\d+)\\).*", "\\1", CIP)) %>%
#     pivot_longer(cols = c(`2019`, `2020`, `2021`), names_to = "Year", values_to = "value") %>%
#     mutate(Year = as.numeric(Year))  # Convert Year column to numeric
#   
#   # Call function to plot line graphs with adjusted y-axis breaks
#   plot_line_graphs(df_long, paste0("OhioJobs_", region, "_Region"))
# }
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# # 
# # # Define a custom color palette with enough colors
# # palette <- c("#E41A1C", "#377EB8", "#4DAF4A", "#984EA3", "#FF7F00", "#FFFF33", "#A65628", "#F781BF", "#999999", "#66C2A5", "#FC8D62", "#8DA0CB", "#E78AC3", "#A6D854", "#FFD92F", "#E5C494", "#B3B3CC", "#8DD3C7", "#FFFFB3", "#BEBADA", "#FB8072", "#80B1D3", "#FDB462", "#B3DE69", "#FCCDE5", "#D9D9D9", "#BC80BD", "#CCEBC5", "#FFED6F")
# # 
# # # Function to plot line graphs
# # plot_line_graphs <- function(data, region) {
# #   # Convert Year column to numeric
# #   data$Year <- as.numeric(data$Year)
# #   
# #   # Create a line plot with points
# #   p <- ggplot(data, aes(x = Year, y = value, group = CIP, color = CIP)) +
# #     geom_line() +
# #     geom_point() +
# #     geom_text(data = subset(data, Year == max(Year)), aes(label = CIP),
# #               nudge_x = 0.1, hjust = 0) +
# #     labs(x = "Year", y = "Graduates", title = region) +
# #     theme_bw() +
# #     theme(legend.position = "right",
# #           legend.background = element_rect(fill = "white", color = "black"),
# #           legend.key = element_rect(color = "black"),
# #           legend.text = element_text(size = 8),
# #           legend.title = element_blank(),
# #           legend.spacing.x = unit(0.2, "cm"),
# #           legend.key.width = unit(0.5, "cm"),  # Adjust the legend key width
# #           legend.key.height = unit(0.5, "cm")) +  # Adjust the legend key height
# #     scale_color_manual(values = palette)  # Use custom color palette
# #   
# #   # Save the plot as an image file
# #   ggsave(paste0(region, "_line_plot.png"), plot = p, width = 10, height = 6)
# # }
# # 
# # # Iterate over unique regions and plot line graphs
# # for (region in unique_regions) {
# #   region_data <- subset(df, JobsOhio.Region == region)
# #   
# #   # Reshape the data to long format for plotting
# #   df_long <- region_data %>%
# #     select(CIP, X2019, X2020, X2021) %>%
# #     rename(`2019` = X2019, `2020` = X2020, `2021` = X2021) %>%
# #     mutate(CIP = gsub(".*\\((\\d+\\.\\d+)\\).*", "\\1", CIP)) %>%
# #     pivot_longer(cols = c(`2019`, `2020`, `2021`), names_to = "Year", values_to = "value") %>%
# #     mutate(Year = as.numeric(Year))  # Convert Year column to numeric
# #   
# #   # Call function to plot line graphs
# #   plot_line_graphs(df_long, region)
# # }
# # 
# # 
# # 
# # 
# # ##PLOTS
# # 
# # # Define a custom color palette with enough colors
# # palette <- c("#E41A1C", "#377EB8", "#4DAF4A", "#984EA3", "#FF7F00", "#FFFF33", "#A65628", "#F781BF", "#999999", "#66C2A5", "#FC8D62", "#8DA0CB", "#E78AC3", "#A6D854", "#FFD92F", "#E5C494", "#B3B3CC", "#8DD3C7", "#FFFFB3", "#BEBADA", "#FB8072", "#80B1D3", "#FDB462", "#B3DE69", "#FCCDE5", "#D9D9D9", "#BC80BD", "#CCEBC5", "#FFED6F")
# # 
# # # Function to plot line graphs
# # plot_line_graphs <- function(data, region) {
# #   # Convert Year column to numeric
# #   data$Year <- as.numeric(data$Year)
# #   
# #   # Create a line plot with points
# #   p <- ggplot(data, aes(x = Year, y = value, group = CIP, color = CIP)) +
# #     geom_line() +
# #     geom_point() +
# #     geom_text(data = subset(data, Year == max(Year)), aes(label = CIP),
# #               nudge_x = 0.1, hjust = 0) +
# #     labs(x = "Year", y = "Graduates", title = region) +
# #     theme_bw() +
# #     theme(legend.position = "bottom",
# #           legend.background = element_rect(fill = "white", color = "black"),
# #           legend.key = element_rect(color = "black"),
# #           legend.text = element_text(size = 8),
# #           legend.title = element_blank(),
# #           legend.spacing.x = unit(0.2, "cm")) +
# #     scale_color_manual(values = palette)  # Use custom color palette
# #   
# #   # Convert the ggplot object to a rvg object
# #   p_dml <- dml(ggobj = p)
# #   
# #   # Create a PowerPoint slide
# #   slide <- read_pptx() %>%
# #     add_slide() %>%
# #     ph_with(p_dml, ph_location())
# #   
# #   # Save the PowerPoint file
# #   print(slide, target = paste0(region, "_line_plot.pptx"))
# # }
# # 
# # # Iterate over unique regions and plot line graphs
# # for (region in unique_regions) {
# #   region_data <- subset(df, JobsOhio.Region == region)
# #   
# #   # Reshape the data to long format for plotting
# #   df_long <- region_data %>%
# #     select(CIP, X2019, X2020, X2021) %>%
# #     rename(`2019` = X2019, `2020` = X2020, `2021` = X2021) %>%
# #     mutate(CIP = gsub(".*\\((\\d+\\.\\d+)\\).*", "\\1", CIP)) %>%
# #     pivot_longer(cols = c(`2019`, `2020`, `2021`), names_to = "Year", values_to = "value") %>%
# #     mutate(Year = as.numeric(Year))  # Convert Year column to numeric
# #   
# #   # Call function to plot line graphs and save to PowerPoint
# #   plot_line_graphs(df_long, paste0("OhioJobs_", region, "_Region"))
# # }
# # 
# 
# 
# 
# # library(ggplot2)
# # 
# # # Filter unique regions
# # unique_regions <- unique(df$JobsOhio.Region)
# # 
# # # Iterate over unique regions and create line charts
# # for (region in unique_regions) {
# #   region_data <- subset(df, JobsOhio.Region == region)
# #   
# #   # Create line chart
# #   ggplot(region_data, aes(x = c("2019", "2020", "2021"), y = c(X2019, X2020, X2021))) +
# #     geom_line() +
# #     labs(x = "Year", y = "Value", title = paste("Line Chart for", region)) +
# #     theme_minimal()
# # }
# # 
# # 
# # 
# write.csv(df, file = paste0(file.path,"cips_organized.csv"), row.names = F)
# 
# # Create a function to generate a line chart and export it to a PowerPoint slide
# create_line_chart_slide <- function(data, region) {
#   # Convert columns to numeric
#   data$X2019 <- as.numeric(data$X2019)
#   data$X2020 <- as.numeric(data$X2020)
#   data$X2021 <- as.numeric(data$X2021)
#   
#   # Create line chart
#   p <- ggplot(data, aes(x = c("2019", "2020", "2021"), y = c(X2019, X2020, X2021))) +
#     geom_line() +
#     labs(x = "Year", y = "Value", title = paste("Line Chart for", region)) +
#     theme_minimal()
#   
#   # Create a PowerPoint object for the line chart
#   p_dml <- dml(ggobj = p)
#   
#   # Create a new PowerPoint slide
#   slide <- read_pptx() %>%
#     add_slide() %>%
#     ph_with(p_dml, ph_location())
#   
#   # Export the slide to a PowerPoint file
#   print(slide, target = "demo.pptx", target_type = "pptx")
# }
# 
# # Filter unique regions
# unique_regions <- unique(df$JobsOhio.Region)
# 
# # Create a PowerPoint presentation
# ppt <- read_pptx()
# 
# # Iterate over unique regions and create line charts on separate slides
# for (region in unique_regions) {
#   region_data <- subset(df, JobsOhio.Region == region)
#   
#   # Create line chart slide for the region
#   create_line_chart_slide(region_data, region)
# }
# 
# # Save the PowerPoint presentation
# print(ppt, target = "demo.pptx")
# 
# 
# 
# 
# 
# "NW"   
# "Regional Growht Partnership"
# 
# 
# "SW" 
# "REDI CINCINNATI"
# 
# "NE" 
# "Team neo"
# 
# "C" 
# "Columbus"
# 
# "W"
# "Daytondevlopment coalition" 
# 
# "SE"
# "Appalachian Partnership"
# Install and load required packages
# install.packages(c("ggplot2", "dplyr", "tidyr", "ggthemes"))
library(ggplot2)
library(dplyr)
library(tidyr)
library(ggthemes)
library(officer)
library(rvg)

file_path <- "H://BB5G/Main_table/attachments/bb_degrees__across.csv"
df <- read.csv(file_path)

# df <- df %>%
#   mutate(jobsohio_region = recode(jobsohio_region,
#                                    "NW" ="Regional Growth Partnership",
#                                    "SW" = "REDI Cincinnati",
#                                    "NE"= "Team NEO",
#                                    "C" =  "OneColumbus",
#                                    "W" = "Dayton Development Coalition",
#                                    "SE"= "Appalachian Partnership"))


# Remove rows with missing values and convert numeric columns to numeric type
df <- df[complete.cases(df), ]
df[, 3:7] <- lapply(df[, 3:7], as.numeric)
# Reorder bars from largest to smallest for each year
# df_ordered <- df %>%
#   select(jobsohio_region, X2019, X2020, X2021) %>%
#   rename(`Year:2019` = X2019,
#          `Year:2020` = X2020,
#          `Year:2021` = X2021) %>%
#   pivot_longer(cols = c(`Year:2019`, `Year:2020`, `Year:2021`), names_to = "Year", values_to = "Number_of_graduates") %>%
#   arrange(Year, desc(Number_of_graduates)) %>%
#   group_by(jobsohio_region, Year) %>%
#   summarize(Number_of_graduates = sum(Number_of_graduates, na.rm = TRUE), .groups = "drop")%>%
#   ungroup()

# Reorder bars from largest to smallest for each year
df_ordered <- df %>%
  dplyr::select(jobsohio_region, X2019, X2020, X2021) %>%
  rename(`Year:2019` = X2019, `Year:2020` = X2020, `Year:2021` = X2021) %>%
  pivot_longer(cols = c(`Year:2019`, `Year:2020`, `Year:2021`), names_to = "Year", values_to = "Number_of_graduates") %>%
  arrange(desc(Number_of_graduates), Year) %>%
  mutate(jobsohio_region = factor(jobsohio_region, levels = c(setdiff(unique(jobsohio_region), "Ohio"), "Ohio"))) %>%
  group_by(jobsohio_region, Year) %>%
  summarize(Number_of_graduates = sum(Number_of_graduates, na.rm = TRUE), .groups = "drop")

# Define custom colors for the bars
bar_colors <- c("#5DA5B3", "#CD5334", "#E0D74E")

# Create a bar graph with grouped bars for each Jobs Ohio region and clustered by Year
p <- ggplot(df_ordered, aes(x = jobsohio_region, y = Number_of_graduates, fill = Year)) +
  geom_bar(stat = "identity", position = position_dodge(width = 0.9), width = 0.7) +
  geom_text(
    aes(label = ifelse(Number_of_graduates == max(Number_of_graduates), Number_of_graduates, "")),
    position = position_dodge(width = 0.9),
    vjust = -0.5,
    color = "black",
    size = 3.5
  ) +
  scale_fill_manual(values = bar_colors) +
  labs(title = "Total Graduates in Ohio by Region", x = "Jobs Ohio Region", y = "Number of Graduates") +
  theme_few() +
  theme(
    plot.title = element_text(size = 16, face = "bold"),
    axis.title.x = element_text(size = 14, family = "Times New Roman"),
    axis.title.y = element_text(size = 14, family = "Times New Roman"),
    axis.text.x = element_text(size = 12, family = "Times New Roman", angle = 90, hjust = 1),
    axis.text.y = element_text(size = 12, family = "Times New Roman"),
    panel.grid.major.y = element_line(color = "gray", linetype = "dashed"),
    panel.border = element_blank(),
    panel.background = element_blank()
  )

# Print the bar chart
print(p)

# Create a PowerPoint object
p_dml <- rvg::dml(ggobj = p)
# Then export the dml object to a PowerPoint slide with officer.

# initialize PowerPoint slide ----
officer::read_pptx() %>%
  # add slide ----
officer::add_slide() %>%
  # specify object and location of object ----
officer::ph_with(p_dml, ph_location()) %>%
  # export slide -----
base::print(
  target = here::here(
    "demo_two.pptx"
  )
)


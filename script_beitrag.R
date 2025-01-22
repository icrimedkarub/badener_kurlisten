# Load necessary libraries
library(tidyverse)
library(lubridate)
library(readxl)
library(dplyr)
library(ggplot2)
library(forcats)
library(scales)

# Load the data
file_path <- "/Users/heima/Desktop/Masterarbeit/Kurlisten/Kurlisten_Excel/kurdata.xlsx"
data <- read_excel(file_path)

file_path <- "/Users/heima/Desktop/Masterarbeit/Kurlisten/Kurlisten_Excel/comparison_data.xlsx"
comparison_data <- read_excel(file_path)

# Convert the 'Date' column to Date format
data <- data %>%
  mutate(Date = dmy(Date))

# Extract the year from the 'Date' column
data <- data %>%
  mutate(Year = year(Date))

# Filter data for the years 1850-1859
filtered_data <- data %>%
  filter(Year >= 1850 & Year <= 1859)

# Summarize the number of visitors by year and sex
summary_data <- filtered_data %>%
  group_by(Year, Sex) %>%
  summarise(Visitors = n(), .groups = "drop")

# Filter out rows with NA values in the 'Sex' column and the specified year range
summary_data <- data %>%
  filter(Year >= 1850 & Year <= 1859, !is.na(Sex)) %>%
  group_by(Year, Sex) %>%
  summarise(Visitors = n(), .groups = "drop")

# Define a common theme for consistent text sizes across plots
common_theme <- theme_minimal() +
  theme(
    axis.text.x = element_text(size = 14), # Common x-axis text size
    axis.text.y = element_text(size = 14), # Common y-axis text size
    axis.title.x = element_text(size = 16), # Common x-axis title text size
    axis.title.y = element_text(size = 16), # Common y-axis title text size
    plot.title = element_text(size = 18, face = "bold", hjust = 0.5), # Common title text size, bold, centered
    legend.title = element_text(size = 16), # Common legend title size
    legend.text = element_text(size = 14), # Common legend text size
    panel.grid.major.x = element_line(color = "gray80", linetype = "solid"), # Light vertical lines
    panel.grid.minor.x = element_blank(), # Remove minor vertical gridlines
    panel.grid.major.y = element_line(color = "gray80", linetype = "dotted"), # Subtle horizontal lines
    panel.grid.minor.y = element_blank() # Remove minor horizontal gridlines
  )

################# TOTAL VISITORS
# Calculate totals for each source
total_offizielleBesucher <- sum(comparison_data$OffizielleBesucher, na.rm = TRUE)
total_Besucher <- sum(comparison_data$Besucher, na.rm = TRUE)

# Define legend labels with totals
label_offizielle <- paste0("offizielle Besucher (", total_offizielleBesucher, ")")
label_Besucher <- paste0("Besucher (", total_Besucher, ")")

# Create the plot with updated legend labels
plot <- ggplot(data = comparison_data, aes(x = Jahr)) +
  geom_vline(
    xintercept = seq(1850, 1859, 1),
    color = "gray80",
    linetype = "solid",
    size = 0.5
  ) +
  geom_line(aes(y = OffizielleBesucher, color = label_offizielle), size = 1) +
  geom_line(aes(y = Besucher, color = label_Besucher), size = 1, linetype = "dashed") +
  geom_point(aes(y = OffizielleBesucher, color = label_offizielle), size = 2) +
  geom_point(aes(y = Besucher, color = label_Besucher), size = 2) +
  scale_color_manual(
    values = setNames(c("blue", "red"), c(label_offizielle, label_Besucher))
  ) +
  scale_x_continuous(
    breaks = seq(1850, 1859, 1),
    labels = scales::number_format(accuracy = 1, big.mark = "")
  ) +
  scale_y_continuous(
    limits = c(6000, 9000),
    breaks = seq(6000, 9000, 1000),
    labels = scales::number_format(accuracy = 1, big.mark = "")
  ) +
  labs(
    title = "Vergleich mit den offiziellen Besucherzahlen",
    x = "Jahr",
    y = "Anzahl der Besucher",
    color = "Quelle"
  ) +
  common_theme

# Display the plot
print(plot)
ggsave("abbildung3b.tiff", plot = plot, dpi = 300, width = 12, height = 8, units = "in", type = "cairo", bg = "transparent")
ggsave("abbildung3c.tiff", plot = plot, dpi = 300, width = 12, height = 8, units = "in", type = "cairo", bg = "white")

################# VISITORS SEX
# Calculate total visitors by sex for the legend
total_visitors <- summary_data %>%
  group_by(Sex) %>%
  summarise(Total = sum(Visitors, na.rm = TRUE)) %>%
  ungroup() # Ensure there is no residual grouping

# Create dynamic labels for the legend including total visitors per sex
legend_labels <- setNames(
  paste0(total_visitors$Sex, " (", total_visitors$Total, ")"),
  total_visitors$Sex
)

# Create the gender balance graph with updated legend
plot <- ggplot(summary_data, aes(x = Year, y = Visitors, color = Sex, group = Sex)) +
  geom_line(linewidth = 1) +
  geom_point(size = 2) +
  scale_color_manual(
    values = c("M" = "blue", "F" = "red"),  # Assign colors
    labels = legend_labels                         # Use dynamic labels with totals
  ) +
  labs(
    title = "Besucherzahlen nach Geschlecht (1850–1859)",
    x = "Jahr",
    y = "Anzahl der Besucher",
    color = "Geschlecht"
  ) +
  scale_x_continuous(
    breaks = seq(1850, 1859, by = 1)
  ) +
  scale_y_continuous(
    breaks = seq(900, 1700, by = 200)
  ) +
  coord_cartesian(ylim = c(900, 1700)) +
  common_theme

# Display the plot
print(plot)
ggsave("abbildung4b.tiff", plot = plot, dpi = 300, width = 12, height = 8, units = "in", type = "cairo", bg = "transparent")
ggsave("abbildung4c.tiff", plot = plot, dpi = 300, width = 12, height = 8, units = "in", type = "cairo", bg = "white")

################# TOTAL IN SEASON PER YEAR
# Assuming data has been loaded and Date column conversion has been attempted
monthly_visitors_by_year <- data %>%
  mutate(
    Month = month(Date, label = TRUE, abbr = TRUE), # Extract month names as abbreviated strings
    Year = year(Date)  # Extract year using lubridate's year() function directly
  ) %>%
  mutate(Month = recode(Month, "May" = "Mai")) %>%  # Change "May" to "Mai"
  filter(!is.na(Month), Month %in% c("Apr", "Mai", "Jun", "Jul", "Aug", "Sep")) %>% # Filter for high season months
  group_by(Year, Month) %>%
  summarise(Total_Visitors = sum(Party, na.rm = TRUE), .groups = "drop")

# Check the data structure again
print(monthly_visitors_by_year)
print(head(data$Date))

# Define color scheme for years
colors <- c("#66c2a5", "#fc8d62", "#8da0cb", "#ffd92f", "#a6d854", "#e78ac3", "#e5c494", "#b3b3b3", "#1b9e77", "#d95f02")

# Create a bar plot showing monthly visitor numbers by year, with a color scheme for years
plot <- ggplot(monthly_visitors_by_year, aes(x = Month, y = Total_Visitors, fill = as.factor(Year))) +
  geom_bar(stat = "identity", position = "dodge", width = 0.7) +  # Position bars side by side for each month
  scale_fill_manual(values = colors) +  # Set color scheme for the years
  labs(
    title = "Gesamtanzahl der Besucher (April bis September 1850–1859) nach Jahr",
    x = "Monat",
    y = "Anzahl der Besucher",
    fill = "Jahr"
  ) +
  scale_y_continuous(
    limits = c(0, 2550),
    breaks = seq(0, 2500, by = 500)
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(size = 18, vjust = 4.5, margin = margin(t = 10)),  # Adjust text spacing as needed
    axis.text.y = element_text(size = 16),
    axis.title.x = element_text(size = 18, vjust = 3),
    axis.title.y = element_text(size = 18),
    plot.title = element_text(size = 18, face = "bold", hjust = 0.5, vjust = -3),
    legend.title = element_text(size = 16),
    legend.text = element_text(size = 14),
    panel.grid.major.x = element_blank(),  # Remove major vertical gridlines
    panel.grid.minor.x = element_blank(),  # Remove minor vertical gridlines
    panel.grid.major.y = element_line(color = "gray80", linetype = "dotted")  # Keep subtle horizontal lines
  )

# Display the plot
print(plot)
ggsave("abbildung5b.tiff", plot = plot, dpi = 300, width = 12, height = 8, units = "in", type = "cairo", bg = "transparent")
ggsave("abbildung5c.tiff", plot = plot, dpi = 300, width = 12, height = 8, units = "in", type = "cairo", bg = "white")

################# TOTAL VISITORS PER COUNTRY EXCL. AUSTRIA
# Calculate total visitors by Country, excluding rows where Country is "Österreich"
top_countries <- data %>%
  filter(Country != "Österreich") %>%
  group_by(Country) %>%
  summarise(Total_Visitors = sum(Party, na.rm = TRUE), .groups = "drop") %>%
  arrange(desc(Total_Visitors)) %>%
  slice_head(n = 10) # Get the top 10 countries

# Create a bar plot for the top 10 countries
plot <- ggplot(top_countries, aes(x = reorder(Country, -Total_Visitors), y = Total_Visitors)) +
  geom_bar(stat = "identity", fill = "darkblue", width = 0.7) +  # Use a single color for all bars
  labs(
    title = "Top 10 Länder nach Besucherzahl (1850–1859) ohne Österreich",
    x = "Land",
    y = "Anzahl der Besucher"
  ) +
  scale_y_continuous(
    limits = c(0, max(top_countries$Total_Visitors, na.rm = TRUE) + 500),
    breaks = scales::pretty_breaks(n = 5)
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(size = 16, angle = 45, hjust = 1), # Rotate and size x-axis labels
    axis.text.y = element_text(size = 16), # Larger y-axis text
    axis.title.x = element_text(size = 16, vjust = 2.5), # Larger x-axis title
    axis.title.y = element_text(size = 16), # Larger y-axis title
    plot.title = element_text(size = 18, face = "bold", hjust = 0.5, vjust = -3), # Bold and larger title
    panel.grid.major.x = element_blank(), # Remove vertical gridlines
    panel.grid.major.y = element_line(color = "gray80", linetype = "dotted") # Subtle horizontal gridlines
  )

# Display the plot
print(plot)
ggsave("abbildung6b.tiff", plot = plot, dpi = 300, width = 12, height = 8, units = "in", type = "cairo", bg = "transparent")
ggsave("abbildung6c.tiff", plot = plot, dpi = 300, width = 12, height = 8, units = "in", type = "cairo", bg = "white")

################# TOTAL VISITORS PER PLACE OVER TIME 5 WITH NUMBERS
# Calculate total visitors per place over all years
places_over_time <- data %>%
  filter(!is.na(Place) & Place != "Wien") %>%
  group_by(Place, Year) %>%
  summarise(Total_Visitors = sum(Party, na.rm = TRUE), .groups = "drop")

# Identify the top 5 places based on total visitors across all years
top_5_places <- places_over_time %>%
  group_by(Place) %>%
  summarise(Total_Visitors_Sum = sum(Total_Visitors, na.rm = TRUE), .groups = "drop") %>%
  arrange(desc(Total_Visitors_Sum)) %>%
  slice_head(n = 5) %>%
  pull(Place)

# Filter data to include only the top 5 places
filtered_places_over_time <- places_over_time %>%
  filter(Place %in% top_5_places)

# Calculate total number of visitors for each place for the legend
total_visitors_per_place <- data %>%
  filter(!is.na(Place) & Place != "Wien") %>%
  group_by(Place) %>%
  summarise(Total_Visitors_Sum = sum(Party, na.rm = TRUE), .groups = "drop")

# Join the total visitors data to the filtered places data
filtered_places_over_time <- filtered_places_over_time %>%
  left_join(total_visitors_per_place, by = "Place") %>%
  mutate(
    Legend_Label = paste(Place, "(", Total_Visitors_Sum, ")", sep = "")
  )

# Define a new color palette
new_colors <- c("blue", "red", "lightgreen", "#984ea3", "orange")

# Plot the data
plot <- ggplot(filtered_places_over_time, aes(x = Year, y = Total_Visitors, color = Legend_Label, group = Place)) +
  geom_line(size = 1) +
  geom_point(size = 2) +
  labs(
    title = "Top 5 Orte nach Besucherzahl (1850–1859) ohne Wien",
    x = "Jahr",
    y = "Anzahl der Besucher",
    color = "Ort"
  ) +
  scale_x_continuous(
    breaks = seq(min(filtered_places_over_time$Year, na.rm = TRUE), max(filtered_places_over_time$Year, na.rm = TRUE), by = 1)
  ) +
  scale_y_continuous(
    limits = c(0, max(filtered_places_over_time$Total_Visitors, na.rm = TRUE) + 1),
    breaks = scales::pretty_breaks(n = 5)
  ) +
  scale_color_manual(values = new_colors) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(size = 14),
    axis.text.y = element_text(size = 14),
    axis.title.x = element_text(size = 16),
    axis.title.y = element_text(size = 16),
    plot.title = element_text(size = 18, face = "bold", hjust = 0.5),
    legend.title = element_text(size = 16),
    legend.text = element_text(size = 14),
    panel.grid.major.x = element_line(color = "gray80", linetype = "solid"),
    panel.grid.minor.x = element_blank(),
    panel.grid.major.y = element_line(color = "gray80", linetype = "dotted")
  )

# Display the plot
print(plot)
ggsave("abbildung7b.tiff", plot = plot, dpi = 300, width = 12, height = 8, units = "in", type = "cairo", bg = "transparent")
ggsave("abbildung7c.tiff", plot = plot, dpi = 300, width = 12, height = 8, units = "in", type = "cairo", bg = "white")

######################  OCCUPATION WITH PARTY
# Calculate the sum of Party values for each occupation category (excluding NA)
occupation_counts <- data %>%
  filter(!is.na(Normalized_Categorized_Occupation)) %>%
  group_by(Normalized_Categorized_Occupation) %>%
  summarise(Count = sum(Party, na.rm = TRUE), .groups = "drop") %>%
  arrange(desc(Count))

# Calculate total number of visitors with occupation info available
total_visitors_with_info <- sum(occupation_counts$Count, na.rm = TRUE)

# Modify the x-axis to show fewer breaks based on the data range
max_count <- max(occupation_counts$Count, na.rm = TRUE)
break_interval <- ceiling(max_count / 5)  # Adjust interval to ensure readability

# Create the bar plot with specified x-axis limits and breaks, and reverse the order of bars
plot <- ggplot(occupation_counts, aes(x = Count, y = fct_rev(fct_reorder(Normalized_Categorized_Occupation, Count)), fill = Normalized_Categorized_Occupation)) +
  geom_bar(stat = "identity", fill = "darkblue") +
  labs(
    title = paste("Besucherzahl nach Berufsklassifikation 1850–1859\nGesamt: ", sum(occupation_counts$Count)),
    x = "Anzahl",
    y = "Berufskategorie"
  ) +
  scale_x_continuous(
    breaks = seq(0, 15000, by = 1500),  # Set breaks every 1500
    limits = c(0, 15000)  # Set maximum limit to 15000
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(size = 16), # Adjust size for better readability
    axis.text.y = element_text(size = 16), # Adjust size for better readability
    axis.title.x = element_text(size = 18), # Larger x-axis title
    axis.title.y = element_text(size = 18, vjust = -2), # Larger y-axis title
    plot.title = element_text(size = 18, face = "bold", hjust = 0.5), # Bold and larger title
    legend.position = "none", # Remove legend if it's not needed
    panel.grid.major.x = element_line(color = "gray80", linetype = "solid"), # Horizontal gridlines
    panel.grid.minor.x = element_blank() # Remove minor vertical lines
  )

# Display the plot
print(plot)
ggsave("abbildung8b.tiff", plot = plot, dpi = 300, width = 12, height = 8, units = "in", type = "cairo", bg = "transparent")
ggsave("abbildung8c.tiff", plot = plot, dpi = 300, width = 12, height = 8, units = "in", type = "cairo", bg = "white")

###################### STACKED BAR CHART FOR OCCUPATION PERCENTAGES
# Manually define a set of colors, potentially repeating some if necessary
colors <- c("#e6194B", "#3cb44b", "#ffe119", "#4363d8", "#f58231", "#911eb4", "#46f0f0", "#f032e6", "darkgreen", "#fabebe",
            "black", "#e6beff", "#9a6324", "#FFF386", "#800000", "#008080", "beige", "#000075", "red")

# Ensure there are enough colors, repeating if necessary
num_categories <- length(unique(occupation_by_year$Normalized_Categorized_Occupation))
if (num_categories > length(colors)) {
  repeat_times <- ceiling(num_categories / length(colors))
  colors <- rep(colors, repeat_times)
}

# Create the stacked bar chart with adjusted y-axis for percentage increments
plot <- ggplot(occupation_by_year, aes(x = as.factor(Year), y = Percentage, fill = Normalized_Categorized_Occupation)) +
  geom_bar(stat = "identity") +
  scale_fill_manual(values = colors[1:num_categories]) +
  labs(
    title = "Prozentuale Verteilung der Besucher nach Berufskategorie pro Jahr",
    x = "Jahr",
    y = "Prozentsatz der Besucher",
    fill = "Berufskategorie"
  ) +
  scale_y_continuous(
    breaks = seq(0, 100, by = 5),  # Set breaks every 5%
    labels = percent_format(scale = 1)  # Format as percentages
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(size = 16, vjust = 3),
    axis.text.y = element_text(size = 16),
    axis.title.x = element_text(size = 18, vjust = 2),
    axis.title.y = element_text(size = 18, vjust = -1),
    plot.title = element_text(size = 18, face = "bold", hjust = 0.5, vjust = -1.5),
    legend.title = element_text(size = 18),
    legend.text = element_text(size = 16),
    panel.grid.major.x = element_blank(),  # Remove major vertical grid lines
    panel.grid.minor.x = element_blank(),  # Remove minor vertical grid lines
    panel.grid.major.y = element_line(color = "gray80"),
    panel.grid.minor.y = element_blank()  # Remove minor horizontal grid lines
  )

# Display the plot
print(plot)
ggsave("abbildung9b.tiff", plot = plot, dpi = 300, width = 12, height = 8, units = "in", type = "cairo", bg = "transparent")
ggsave("abbildung9c.tiff", plot = plot, dpi = 300, width = 12, height = 8, units = "in", type = "cairo", bg = "white")


###################### AVERAGE PARTY SIZE PER CATEGORY
# Calculate and sort the average party size for each occupation category
average_party_size_by_occupation <- data %>%
  filter(!is.na(Normalized_Categorized_Occupation)) %>%  # Exclude rows where the occupation category is NA
  group_by(Normalized_Categorized_Occupation) %>%
  summarise(
    Average_Party_Size = mean(Party, na.rm = TRUE),  # Calculate mean, ignoring NA values in 'Party'
    .groups = 'drop'  # Remove grouping structure from the resulting tibble
  ) %>%
  arrange(desc(Average_Party_Size))  # Order the results by average party size in descending order

# View the results
print(average_party_size_by_occupation)

###################### TOP 5 WITH COLUMN
# Calculate the frequency of each entry in the 'With' column and find the top 5
top_five_with <- data %>%
  filter(!is.na(With)) %>%  # Exclude NA values to focus on valid entries
  count(With, sort = TRUE) %>%  # Count occurrences and sort in descending order
  top_n(5, n)  # Select the top 5 entries

# View the results
print(top_five_with)
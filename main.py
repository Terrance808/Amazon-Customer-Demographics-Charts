import pandas as pd
import matplotlib.pyplot as plt


def main():
    # Read Excel sheet and format dataframe columns
    df = pd.read_excel("Amazon-Demographics.xlsx", sheet_name="Second Half 2022")
    df["Age"] = df["Age"].astype(str)
    df["Education"] = df["Education"].astype(str)
    df["Gender"] = df["Gender"].astype(str)
    df["Marital Status"] = df["Marital Status"].astype(str)

    # Create and format subplots for first visual
    fig, axes = plt.subplots(3, 1, figsize=(12.8, 8.6),  # 3 row charts
                             height_ratios=[2 / 10, 3 / 10, 3 / 10],
                             gridspec_kw={"left": 1 / 10, "right": 9 / 10, "hspace": 0.5})
    fig.suptitle("Amazon Customer Demographics", fontsize=24)

    # Chart 1, first visual
    ax1 = axes[0]
    ax1.bar(df["Age"], df["Age Count"], width=-0.4, align="edge", color="tab:blue")
    ax1.set_xlabel("Age")
    ax1.set_ylabel("Count")
    ax1_xticks = ax1.get_xticklabels()
    ax1.set_xticklabels(ax1_xticks, ha="right")
    ax1.bar_label(ax1.containers[0], padding=-12)

    # Chart 2, second visual
    ax2 = axes[1]
    ax2.bar(df["Income"], df["Income Count"], width=-0.4, align="edge", color="tab:orange")
    ax2.set_xlabel("Income")
    ax2.set_ylabel("Count")
    ax2_xticks = ax2.get_xticklabels()
    ax2.set_xticklabels(ax2_xticks, rotation=10, ha="center")
    ax2.bar_label(ax2.containers[0], padding=-12)

    # Chart 3, third visual
    ax3 = axes[2]
    ax3.bar(df["Education"], df["Education Count"], width=-0.4, align="edge", color="tab:blue")
    ax3.set_xlabel("Education")
    ax3.set_ylabel("Count")
    ax3_xticks = ax3.get_xticklabels()
    ax3.set_xticklabels(ax3_xticks, ha="right")
    ax3.bar_label(ax3.containers[0], padding=-12)

    plt.show()
    plt.close()

    # Create and format subplots for second visual
    fig, axes = plt.subplots(1, 2, figsize=(12.8, 8.6))
    fig.suptitle("Amazon Customer Demographics", fontsize=24)

    # Chart 1, second visual
    ax1 = axes[0]
    ax1.bar(df["Gender"], df["Gender Count"], width=-0.4, align="edge", color="tab:blue")
    ax1.set_xlabel("Gender")
    ax1.set_ylabel("Count")
    ax1_xticks = ax1.get_xticklabels()
    ax1.set_xticklabels(ax1_xticks, ha="right")
    ax1.bar_label(ax1.containers[0], padding=-12)

    # Chart 2, second visual
    ax2 = axes[1]
    ax2.bar(df["Marital Status"], df["Marital Status Count"], width=-0.4, align="edge", color="tab:orange")
    ax2.set_xlabel("Marital Status")
    ax2.set_ylabel("Count")
    ax2_xticks = ax2.get_xticklabels()
    ax2.set_xticklabels(ax2_xticks, ha="right")
    ax2.bar_label(ax2.containers[0], padding=-12)

    plt.show()
    plt.close()


if __name__ == "__main__":
    main()
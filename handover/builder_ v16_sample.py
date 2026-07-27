# ------------------------------------------------------------
# デジタル庁 公式カラーパレット（Phase A-1）
# 出典：デジタル庁 ダッシュボードデザインの実践ガイドブックとデザインテンプレート
# https://www.digital.go.jp/resources/dashboard-guidebook
# ライセンス：PDL1.0（公共データ利用規約 第1.0版）
# GitHub：https://github.com/digital-go-jp/policy-dashboard-assets
# 取得日：2026-07-27
# ------------------------------------------------------------
DIGITAL_AGENCY_PALETTE = {
    "SolidGray": {
        "primary":  "#4D4D4D",
        "secondary":"#767676",
        "midtone":  "#999999",
        "light":    "#CCCCCC",
        "lightest": "#F2F2F2",
        "accent":   "#3460FB",  # 強調用（青）
        "warning":  "#FE3939",  # 警告用（赤）
        "bg":       "#F8F8FB",
    },
    "Blue": {
        "primary":  "#0017C1",
        "secondary":"#3460FB",
        "midtone":  "#7096F8",
        "light":    "#C5D7FB",
        "lightest": "#E8F1FE",
        "accent":   "#FE3939",
        "warning":  "#FFBBBB",
        "bg":       "#F8F8FB",
    },
    "LightBlue": {
        "primary":  "#0055AD",
        "secondary":"#008BF2",
        "midtone":  "#57B8FF",
        "light":    "#C0E4FF",
        "lightest": "#F0F9FF",
        "accent":   "#FE3939",
        "warning":  "#FFBBBB",
        "bg":       "#F8F8FB",
    },
    "Green": {
        "primary":  "#115A36",
        "secondary":"#259D63",
        "midtone":  "#51B883",
        "light":    "#9BD4B5",
        "lightest": "#E6F5EC",
        "accent":   "#666666",
        "warning":  "#CCCCCC",
        "bg":       "#F8F8FB",
    },
    "Cyan": {
        "primary":  "#006F83",
        "secondary":"#00A3BF",
        "midtone":  "#2BC8E4",
        "light":    "#99F2FF",
        "lightest": "#E9F7F9",
        "accent":   "#666666",
        "warning":  "#CCCCCC",
        "bg":       "#F8F8FB",
    },
    "Red": {
        "primary":  "#CE0000",
        "secondary":"#FE3939",
        "midtone":  "#FF7171",
        "light":    "#FFBBBB",
        "lightest": "#FDEEEE",
        "accent":   "#666666",
        "warning":  "#CCCCCC",
        "bg":       "#F8F8FB",
    },
    "Orange": {
        "primary":  "#AC3E00",
        "secondary":"#FB5B01",
        "midtone":  "#FF8D44",
        "light":    "#FFC199",
        "lightest": "#FFEEE2",
        "accent":   "#666666",
        "warning":  "#CCCCCC",
        "bg":       "#F8F8FB",
    },
}

# 全テーマ共通の閾値色（good/bad判定用）
DIGITAL_AGENCY_THRESHOLD = {
    "center":  "#E6E6E6",  # 中央値・中立表示
    # maximum / minimum は各テーマの4番目色に準ずる（テーマ依存）
}

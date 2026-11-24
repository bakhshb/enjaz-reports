"""Header row heights (EMU) shared by the weekly consolidated and sector decks."""

# Title and header row height of the summary task tables in the approved
# masters (frame 1203960 minus the 324198 body row, split over the two rows).
# The masters leave these rows auto-height; طلبات الدعم uses the same value.
# 0.48 in.
TASK_HEADER = 439881

# A header that wraps to two lines ("تاريخ الإنجاز المخطط", "تاريخ البداية",
# "تاريخ انشاء المعاملة") needs 0.68 in at 11 pt bold Abar. The masters' own
# rows (0.52 in projects, 0.557 in transactions, 0.46 in on the digital master)
# clip the second line in the LibreOffice render. Checked by rendering.
TWO_LINE_HEADER = 621792

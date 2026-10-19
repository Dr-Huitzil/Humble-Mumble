"""NinjaOne Inventory & Dashboard Reporting Module."""

from humble_mumble.ninjaone.build_inventory import build_inventory_sheet, main
from humble_mumble.ninjaone.generate_dashboard_tabs import generate_dashboards

__all__ = ["build_inventory_sheet", "generate_dashboards", "main"]

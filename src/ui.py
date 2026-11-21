import streamlit as st
import streamlit.components.v1 as components

TAILWIND_CDN = '<script src="https://cdn.tailwindcss.com"></script>'

def render_header(title: str, subtitle: str, badge_text: str = "Workforce Analytics v1.0"):
    """Renders a modern, responsive corporate header banner using Tailwind CSS."""
    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        {TAILWIND_CDN}
        <style>body {{ margin: 0; padding: 0; background-color: transparent; font-family: system-ui, -apple-system, sans-serif; }}</style>
    </head>
    <body>
        <div class="bg-gradient-to-r from-blue-900 via-blue-800 to-indigo-900 text-white p-6 rounded-2xl shadow-lg border border-blue-700/50 flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
            <div>
                <div class="flex items-center gap-2 mb-2">
                    <span class="bg-blue-600/60 text-blue-200 text-xs font-semibold px-3 py-1 rounded-full border border-blue-400/30 uppercase tracking-wider">{badge_text}</span>
                    <span class="text-blue-300 text-xs">PRJ-AIA-01 | ABES MCA 2nd Yr</span>
                </div>
                <h1 class="text-2xl md:text-3xl font-extrabold tracking-tight text-white">{title}</h1>
                <p class="text-blue-200 text-sm mt-1 max-w-2xl">{subtitle}</p>
            </div>
            <div class="bg-blue-800/40 border border-blue-400/20 px-4 py-2 rounded-xl text-right shrink-0">
                <span class="text-xs text-blue-300 block font-medium">Author / Lead Analyst</span>
                <span class="text-sm font-bold text-white block">Kalpana (2500320140071)</span>
            </div>
        </div>
    </body>
    </html>
    """
    components.html(html_code, height=135)

def render_kpi_card(title: str, value: str, subtitle: str, trend: str = "", color_theme: str = "blue"):
    """Renders a single Tailwind CSS metric card."""
    color_map = {
        "blue": ("border-blue-500", "text-blue-600", "bg-blue-50"),
        "green": ("border-emerald-500", "text-emerald-600", "bg-emerald-50"),
        "red": ("border-rose-500", "text-rose-600", "bg-rose-50"),
        "purple": ("border-purple-500", "text-purple-600", "bg-purple-50"),
        "amber": ("border-amber-500", "text-amber-600", "bg-amber-50")
    }
    border_c, text_c, bg_c = color_map.get(color_theme, color_map["blue"])
    
    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        {TAILWIND_CDN}
        <style>body {{ margin: 0; padding: 0; background-color: transparent; font-family: system-ui, -apple-system, sans-serif; }}</style>
    </head>
    <body>
        <div class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200/80 hover:shadow-md transition-shadow relative overflow-hidden border-l-4 {border_c}">
            <div class="flex items-center justify-between mb-2">
                <span class="text-xs font-semibold uppercase tracking-wider text-slate-500">{title}</span>
                <span class="text-xs font-bold px-2 py-0.5 rounded-md {bg_c} {text_c}">{trend}</span>
            </div>
            <div class="text-3xl font-black text-slate-800 tracking-tight">{value}</div>
            <p class="text-xs text-slate-500 mt-2 font-medium">{subtitle}</p>
        </div>
    </body>
    </html>
    """
    components.html(html_code, height=125)

def render_kpi_grid_4(cards_data):
    """Renders 4 metric cards side by side in one responsive iframe container."""
    card_htmls = ""
    color_map = {
        "blue": ("border-blue-500", "text-blue-600", "bg-blue-50"),
        "green": ("border-emerald-500", "text-emerald-600", "bg-emerald-50"),
        "red": ("border-rose-500", "text-rose-600", "bg-rose-50"),
        "purple": ("border-purple-500", "text-purple-600", "bg-purple-50"),
        "amber": ("border-amber-500", "text-amber-600", "bg-amber-50")
    }
    for title, val, sub, trend, c_theme in cards_data:
        border_c, text_c, bg_c = color_map.get(c_theme, color_map["blue"])
        card_htmls += f"""
        <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200/80 hover:shadow-md transition-shadow border-l-4 {border_c} flex-1 min-w-[220px]">
            <div class="flex items-center justify-between mb-1.5">
                <span class="text-xs font-semibold uppercase tracking-wider text-slate-500">{title}</span>
                <span class="text-xs font-bold px-2 py-0.5 rounded-md {bg_c} {text_c}">{trend}</span>
            </div>
            <div class="text-2xl md:text-3xl font-black text-slate-800 tracking-tight">{val}</div>
            <p class="text-xs text-slate-500 mt-1.5 font-medium">{sub}</p>
        </div>
        """
        
    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        {TAILWIND_CDN}
        <style>body {{ margin: 0; padding: 0; background-color: transparent; font-family: system-ui, -apple-system, sans-serif; }}</style>
    </head>
    <body>
        <div class="flex flex-wrap md:flex-nowrap gap-4 w-full">
            {card_htmls}
        </div>
    </body>
    </html>
    """
    components.html(full_html, height=130)

def render_insight_box(title: str, message: str, alert_type: str = "info"):
    """Renders a styled callout insight box."""
    type_map = {
        "info": ("bg-blue-50", "border-blue-300", "text-blue-900", "text-blue-700"),
        "warning": ("bg-amber-50", "border-amber-300", "text-amber-900", "text-amber-700"),
        "danger": ("bg-rose-50", "border-rose-300", "text-rose-900", "text-rose-700"),
        "success": ("bg-emerald-50", "border-emerald-300", "text-emerald-900", "text-emerald-700")
    }
    bg, border, text_title, text_body = type_map.get(alert_type, type_map["info"])
    
    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        {TAILWIND_CDN}
        <style>body {{ margin: 0; padding: 0; background-color: transparent; font-family: system-ui, -apple-system, sans-serif; }}</style>
    </head>
    <body>
        <div class="{bg} border-l-4 {border} p-4 rounded-r-xl shadow-sm">
            <div class="flex items-center gap-2">
                <span class="font-bold text-sm {text_title}">{title}</span>
            </div>
            <p class="text-xs {text_body} mt-1 leading-relaxed">{message}</p>
        </div>
    </body>
    </html>
    """
    components.html(html_code, height=95)

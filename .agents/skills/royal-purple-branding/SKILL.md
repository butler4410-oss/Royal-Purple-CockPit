---
name: royal-purple-branding
description: Royal Purple brand styling guidelines for Streamlit UI components. Use when building or modifying pages in the Partnership Hub, adding new branded sections, or maintaining consistent visual design across the app.
---

# Royal Purple Branding

Styling patterns and brand guidelines used across the Royal Purple Partnership Hub Streamlit app.

## Brand Colors

| Name | Hex | Usage |
|---|---|---|
| RP Deep Purple | `#2D1B5E` | Gradient dark end, heading text |
| RP Purple | `#4B2D8A` | Primary brand color, buttons, accents |
| RP Light Purple | `#6B3FA0` | Gradient light end |
| RP Lavender | `#C4B5E8` | Subtitle text on dark backgrounds |
| RP Soft Lavender | `#A78BCC` | Secondary text on dark backgrounds |
| RP Background | `#F8F5FF` | Light purple tinted backgrounds |
| RP Border | `#EDE9FE` | Section divider borders |
| Slate Gray | `#94A3B8` | Subtitle text on light backgrounds |
| Body Text | `#4B5563` | Standard body text |

## Page Hero Banner

Use on every main page for a branded header:

```python
st.markdown(
    """
    <div style="background:linear-gradient(135deg,#2D1B5E 0%,#4B2D8A 60%,#6B3FA0 100%);
                border-radius:12px;padding:32px 36px 28px;margin-bottom:8px;">
        <div style="font-size:11px;font-weight:700;letter-spacing:3px;color:#C4B5E8;
                    text-transform:uppercase;margin-bottom:8px;">Royal Purple Partnership Hub</div>
        <div style="font-size:28px;font-weight:800;color:#FFFFFF;line-height:1.2;margin-bottom:8px;">
            Page Title Here
        </div>
        <div style="font-size:14px;color:#C4B5E8;max-width:560px;line-height:1.6;">
            Page description goes here — one or two lines explaining the tool.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)
```

## Section Labels

Small uppercase tracking labels with a purple accent tag:

```python
st.markdown(
    """
    <div style="border-top:2px solid #EDE9FE;margin:8px 0 16px;">
        <span style="display:inline-block;background:#4B2D8A;color:white;
                     font-size:10px;font-weight:700;letter-spacing:2px;
                     text-transform:uppercase;padding:3px 10px;border-radius:0 0 6px 6px;">
            Section Name
        </span>
    </div>
    """,
    unsafe_allow_html=True,
)
```

## Subsection Headers

Lighter, smaller uppercase labels above metric rows:

```python
st.markdown(
    """
    <div style="font-size:10px;font-weight:700;letter-spacing:2.5px;color:#9CA3AF;
                text-transform:uppercase;margin-bottom:4px;">Network Summary</div>
    """,
    unsafe_allow_html=True,
)
```

## Info Cards

Purple-accented cards for contextual information:

```python
st.markdown(
    """
    <div style="background:#F8F5FF;border-left:4px solid #4B2D8A;border-radius:0 8px 8px 0;
                padding:16px 18px;margin-top:8px;">
        <div style="font-weight:700;color:#2D1B5E;font-size:13px;margin-bottom:8px;">
            Card Title
        </div>
        <div style="font-size:12px;color:#4B5563;line-height:1.8;">
            ✦ &nbsp;Bullet point one<br>
            ✦ &nbsp;Bullet point two
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)
```

## Success Messages

Custom styled success banners (instead of plain `st.success`):

```python
st.markdown(
    f"""
    <div style="background:#F0FDF4;border:1px solid #86EFAC;border-radius:8px;
                padding:12px 18px;margin:12px 0 20px;">
        <span style="color:#166534;font-weight:700;">
            Main message here
        </span>
        <span style="color:#4B7A5E;font-size:13px;">secondary detail</span>
    </div>
    """,
    unsafe_allow_html=True,
)
```

## Dark CTA Blocks

For prominent call-to-action sections (e.g., report generation):

```python
st.markdown(
    f"""
    <div style="background:linear-gradient(135deg,#1E0F3C 0%,#2D1B5E 100%);
                border-radius:10px;padding:24px 28px;margin:24px 0 8px;">
        <div style="color:#C4B5E8;font-size:11px;font-weight:700;letter-spacing:2px;
                    text-transform:uppercase;margin-bottom:6px;">Ready to Export</div>
        <div style="color:white;font-size:18px;font-weight:700;margin-bottom:4px;">
            Action Title
        </div>
        <div style="color:#A78BCC;font-size:13px;">
            Description of what happens next
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)
st.button("Primary Action", type="primary", use_container_width=True)
```

## Sidebar

The sidebar uses a simple text-only layout (no logo images):

```python
with st.sidebar:
    st.markdown("")
    st.markdown("**The Royal Purple Partnership Hub**")
    st.caption("by ThrottlePro")
    st.markdown("---")
    nav = st.radio("Navigation", [...], label_visibility="collapsed")
    st.markdown("---")
    st.image(LOGO_NEVER_SETTLE, width="stretch")  # Never Settle background only
    st.caption("More Cars. More Loyalty. Less Stress.")
```

## Account Type Colors (Map)

8 account types with assigned colors:

| Type | Color | Hex |
|---|---|---|
| Promo Only | Red | `#DC2626` |
| On Both Lists | Green | `#16A34A` |
| C4C Only | Blue | `#2563EB` |
| Rack Installer | Purple | `#7C3AED` |
| Distributor | Gold | `#F59E0B` |
| Powersports/Motorsports | Rose | `#E11D48` |
| International | Indigo | `#4F46E5` |
| Canada | Emerald | `#059669` |

## General Rules

- Use `unsafe_allow_html=True` for all custom HTML/CSS in `st.markdown()`
- Buttons: use `type="primary"` for main actions, default for secondary
- Use `use_container_width=True` on prominent buttons for full-width
- Metrics: use standard `st.metric()` — they inherit Streamlit's theme
- Use `st.caption()` for helper text beneath sections
- Use `✦` (not emoji) as bullet markers in info cards
- Font stack: inherit Streamlit's default (no custom fonts)
- All styling uses inline CSS (no external stylesheets)

## Asset Paths

```python
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "RP_Synthetic_Expert_Logo_Black_Text.png")
LOGO_SIDEBAR_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "RPMO_logo_BF_Outline.png")
LOGO_NEVER_SETTLE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "25-RYP-02147 Employee LinkedIn Thumbnails P1-6.jpg")
```

**Critical rule:** No logo images in PPTX slides. The Never Settle background image is used only on PPTX cover/divider/closing slides, never on data slides.

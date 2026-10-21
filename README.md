# MYPG 四组晚餐服事 · 手机点餐页

A phone-first, single-page web app that replaces the Google Doc + Google Sheet workflow for a church fellowship group's monthly dinner service (~15 people).

**Live:** https://martinl13.github.io/mypg-dinner/

## What it does

Each month everyone opens one link from the WeChat group and, in under a minute:

1. Taps their own name
2. Picks how they're joining this month — **拼菜** (chip in $25 for restaurant dishes), **自带** (bring a homemade dish), or **不来** (skip)
3. Does only the part that applies to them:
   - 拼菜 → Venmo deep link / Zelle copy button with the note prefilled, then "+1" up to three dishes from the restaurant menu
   - 自带 → registers the dish; duplicates are flagged so nobody brings the same thing twice
   - 不来 → done

Everyone sees the live board: dish ranking by votes, the potluck wall, and how many dishes will be ordered (`chip-in count ÷ 2`, one dish ≈ $50). The organizer view at the bottom shows who hasn't paid and a copy-ready order list for the restaurant.

## Status

**Prototype.** Data is sample data and resets on refresh. Next step is a Google Apps Script backend that reads the menu from, and writes sign-ups back to, the group's existing Google Sheet — so swapping restaurants is just pasting a new menu tab, and historical records stay in one place.

## Stack

Plain HTML/CSS/JS, no build step. Hosted on GitHub Pages. The site never handles money — payment links only.

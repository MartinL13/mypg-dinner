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

**Live (Oct 2026).** The page reads from and writes to the group's existing Google Sheet through a Google Apps Script web app (`apps-script/Code.gs`). Sign-ups land in the same name × date grid the group has always used, so records before and after the switch are continuous. The restaurant menu lives in a `菜单` tab — swapping restaurants is pasting a new menu, no code change.

Setup steps for the Sheet side are in `apps-script/部署说明.md`.

## Stack

Plain HTML/CSS/JS, no build step, hosted on GitHub Pages. Backend is Google Apps Script bound to the Sheet (free, no server). The site never handles money — payment links only.

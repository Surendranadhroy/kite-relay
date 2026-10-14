<!DOCTYPE html>
<html lang="en" data-theme="dark">
<!-- v5 — 7 Oct 2026: Live rows expand + refresh every 30 s, Live order book with note icon beside status and change/cancel for pending orders, pending Mock orders editable, per-user Live auto-trading switch, per-user proxy field, charge breakdown tooltip, per-trade cap = margin. See CHANGES_V5.md. -->
<!-- v4 — 5 Oct 2026: Live tab = Holdings | Positions | Trade Log with the Action column, live orders in the Order Book (Mock | Live), manual live orders from the order form, Guide removed from the Live tab, owner-aware Broker (Live) settings. See CHANGES_V4.md. -->
<!-- v3 — 3 Oct 2026: compact toolbar layout for Portfolio / Watchlist / Live, Live tab with summaries + filters, Broker (Live) setup guide (TrueIP + GitHub + Render, with the relay files), weekend & holiday trading switch, IPFT, icon-only download, MS-Authenticator-safe QR. See CHANGES_V3.md. -->
<!-- v2 — Oct 2026 additions (marked "v2" in the code): two-step verification screens, Brokerage / Broker (Live) / selection-limit settings, charges-aware P&L (FIFO per-trade lines), Holdings+Positions for stocks traded today, category filters, per-category headers, margin strips, Live fixes. See CHANGES_AND_SETUP.md. -->
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no"/>
<title>Signal Performance</title>
<link href="https://fonts.googleapis.com/css2?family=Bebas+Neue&family=IBM+Plex+Mono:wght@300;400;600&family=Barlow+Condensed:wght@300;500;700&display=swap" rel="stylesheet"/>
<script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script><!-- v2: draws the two-step-verification QR code locally; if it fails to load the manual key is still shown -->
<style>
:root,[data-theme="dark"]{--bg:#080b10;--sf:#0e1420;--sf2:#141c2b;--bd:#1e2d42;--gr:#00e68a;--gdm:rgba(0,230,138,.10);--rd:#ff3f5b;--rdm:rgba(255,63,91,.10);--am:#ffa94d;--bl:#4da6ff;--pu:#b48eff;--cy:#00d4ff;--mu:#4a607a;--tx:#c8d8ec;--dm:#6a8099;--sh:0 4px 24px rgba(0,0,0,.6);--sc:rgba(0,0,0,.04);--ac:#00e68a;--ad:rgba(0,230,138,.10);}
[data-theme="light"]{--bg:#edf1f7;--sf:#ffffff;--sf2:#f0f4f9;--bd:#c0cfe0;--gr:#006e40;--gdm:rgba(0,110,64,.10);--rd:#b91c36;--rdm:rgba(185,28,54,.10);--am:#a66200;--bl:#1a54b5;--pu:#5c2fa0;--cy:#007a94;--mu:#64788e;--tx:#111827;--dm:#334d66;--sh:0 2px 12px rgba(0,0,0,.14),0 0 0 1px rgba(0,0,0,.06);--sc:rgba(0,0,0,.06);--ac:#006e40;--ad:rgba(0,110,64,.10);}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
html,body{width:100%;overflow-x:hidden;-webkit-text-size-adjust:100%}
body{background:var(--bg);color:var(--tx);font-family:'IBM Plex Mono',monospace;font-size:13px;min-height:100dvh;-webkit-tap-highlight-color:transparent;transition:background .3s,color .3s;-webkit-user-select:none;user-select:none}
input,textarea{-webkit-user-select:text;user-select:text}
input[type="date"]{color-scheme:dark}
[data-theme="light"] input[type="date"]{color-scheme:light}
#sbT{position:fixed;left:0;top:0;width:6px;height:100%;z-index:601;display:none}
#sbO{display:none;position:fixed;inset:0;z-index:599;background:rgba(0,0,0,.4)}#sbO.show{display:block}
#sb{position:fixed;left:0;top:0;height:100%;width:48px;background:var(--sf);border-right:1px solid var(--bd);z-index:600;transform:translateX(-100%);transition:transform .28s cubic-bezier(.16,1,.3,1);display:flex;flex-direction:column;align-items:center;padding:8px 0;gap:2px;box-shadow:var(--sh)}
#sb.open{transform:translateX(0)}
#aiT{position:fixed;right:0;top:0;width:6px;height:100%;z-index:601;display:none}
#aiO{display:none;position:fixed;inset:0;z-index:589;background:rgba(0,0,0,.4)}#aiO.show{display:block}
#aip{position:fixed;right:0;top:0;height:100%;width:360px;background:var(--sf);border-left:1px solid var(--bd);z-index:590;transform:translateX(100%);transition:transform .28s cubic-bezier(.16,1,.3,1);display:flex;flex-direction:column;box-shadow:var(--sh)}
#aipResize{position:absolute;left:-4px;top:0;bottom:0;width:8px;cursor:ew-resize;z-index:10;touch-action:none}
#aipResize:hover,#aipResize.dragging{background:var(--ac);opacity:.35}
#aip.open{transform:translateX(0)}
.aihd{padding:10px 13px;border-bottom:1px solid var(--bd);background:var(--sf2);display:flex;align-items:center;gap:8px;flex-shrink:0}
.aittl{font-family:'Bebas Neue',sans-serif;font-size:16px;letter-spacing:2px;color:var(--ac);flex:1}
.aisub{font-family:'Barlow Condensed',sans-serif;font-size:9px;letter-spacing:1.5px;color:var(--dm);text-transform:uppercase;flex-shrink:0}
.aibd{flex:1;overflow-y:auto;padding:10px;display:flex;flex-direction:column;gap:8px;-webkit-overflow-scrolling:touch}
.ai-welcome{text-align:center;padding:18px 10px;color:var(--dm);font-family:'Barlow Condensed',sans-serif;font-size:12px;letter-spacing:1px;line-height:1.9}
.ai-title{font-family:'Bebas Neue',sans-serif;font-size:20px;letter-spacing:3px;color:var(--ac);display:block;margin-bottom:8px}
.ai-chips{display:flex;flex-wrap:wrap;gap:5px;margin-top:10px;justify-content:center}
.ai-chip{padding:5px 9px;background:var(--sf2);border:1px solid var(--bd);border-radius:16px;font-family:'Barlow Condensed',sans-serif;font-size:10px;letter-spacing:1px;color:var(--dm);cursor:pointer;text-transform:uppercase;transition:all .2s}
.ai-chip:hover{border-color:var(--ac);color:var(--ac)}
.ai-msg{display:flex;flex-direction:column;gap:3px;animation:fadeUp .25s ease}
.ai-msg.user{align-items:flex-end}.ai-msg.bot{align-items:flex-start}
.ai-bubble{padding:8px 12px;border-radius:10px;font-size:12px;line-height:1.6;max-width:93%;word-wrap:break-word;white-space:pre-wrap}
.ai-msg.bot .ai-bubble{white-space:normal}
.ai-msg.user .ai-bubble{background:var(--ad);border:1px solid var(--ac);color:var(--tx);border-bottom-right-radius:2px}
.ai-msg.bot  .ai-bubble{background:var(--sf2);border:1px solid var(--bd);color:var(--tx);border-bottom-left-radius:2px}
.ai-bubble .ai-p{margin:0 0 8px}
.ai-bubble .ai-p:last-child{margin-bottom:0}
.ai-bubble .ai-h{font-weight:700;color:var(--ac);margin:10px 0 6px;line-height:1.3}
.ai-bubble .ai-h:first-child{margin-top:0}
.ai-bubble .ai-h1{font-size:14px}.ai-bubble .ai-h2{font-size:13px}.ai-bubble .ai-h3,.ai-bubble .ai-h4,.ai-bubble .ai-h5,.ai-bubble .ai-h6{font-size:12px;text-transform:uppercase;letter-spacing:.5px;color:var(--dm)}
.ai-bubble ul,.ai-bubble ol{margin:0 0 8px;padding-left:18px}
.ai-bubble li{margin-bottom:3px}
.ai-bubble code{background:var(--bg);border:1px solid var(--bd);border-radius:3px;padding:1px 4px;font-size:11px;font-family:monospace}
.ai-bubble b{color:var(--tx);font-weight:700}
.ai-tblwrap{overflow-x:auto;margin:0 0 8px;border:1px solid var(--bd);border-radius:6px}
.ai-tbl{width:100%;border-collapse:collapse;font-size:11px;white-space:nowrap}
.ai-tbl th{background:var(--sf);color:var(--dm);text-transform:uppercase;letter-spacing:.5px;font-size:9px;padding:6px 8px;text-align:left;border-bottom:1px solid var(--bd)}
.ai-tbl td{padding:6px 8px;border-bottom:1px solid var(--bd);white-space:normal}
.ai-tbl tr:last-child td{border-bottom:none}
.ai-dots{display:flex;gap:4px;padding:8px 12px;background:var(--sf2);border:1px solid var(--bd);border-radius:10px;border-bottom-left-radius:2px;width:fit-content}
.ai-dots span{width:6px;height:6px;border-radius:50%;background:var(--dm);animation:aiDot 1.2s ease-in-out infinite}
.ai-dots span:nth-child(2){animation-delay:.2s}.ai-dots span:nth-child(3){animation-delay:.4s}
@keyframes aiDot{0%,80%,100%{transform:scale(1);opacity:.4}40%{transform:scale(1.2);opacity:1}}
.ai-inprow{padding:9px;border-top:1px solid var(--bd);background:var(--sf2);display:flex;gap:6px;flex-shrink:0}
.ai-inp{flex:1;padding:7px 11px;background:var(--sf);border:1px solid var(--bd);border-radius:8px;color:var(--tx);font-family:'IBM Plex Mono',monospace;font-size:12px;outline:none;resize:none;line-height:1.4;max-height:80px}
.ai-inp:focus{border-color:var(--ac)}.ai-inp::placeholder{color:var(--dm)}
.ai-send{width:34px;height:34px;background:var(--ac);border:none;border-radius:8px;color:#000;font-size:16px;cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0;align-self:flex-end}
.ai-send:disabled{opacity:.4;cursor:not-allowed}
.sbb,.sba{position:relative;width:36px;height:36px;background:transparent;border:none;cursor:pointer;display:flex;align-items:center;justify-content:center;border-radius:9px;transition:background .18s;flex-shrink:0}
.sbb:hover,.sba:hover{background:var(--sf2)}.sbb.lon{background:var(--ad)}
.sbico{font-size:14px;color:var(--dm)}
.refBox{display:flex;flex-direction:column;align-items:center;gap:2px;padding:7px 2px;cursor:pointer;border-top:1px solid var(--bd);border-bottom:1px solid var(--bd);margin:4px 0;color:var(--dm);transition:color .15s}
.refBox:hover{color:var(--ac)}
.refTxt{font-size:8px;font-family:'Barlow Condensed',sans-serif;letter-spacing:.5px;text-transform:uppercase}
.shareGrid{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.shareBtn{display:flex;align-items:center;gap:8px;padding:10px 12px;background:var(--bg);border:1px solid var(--bd);border-radius:8px;cursor:pointer;font-size:12px;color:var(--tx);transition:border-color .15s}
.shareBtn:hover{border-color:var(--ac)}
.shIco{font-weight:700;font-size:15px;width:18px;text-align:center;flex-shrink:0}
.sbb.lon[data-lp="intraday"] .sbico{color:var(--cy)}.sbb.lon[data-lp="shortterm"] .sbico{color:var(--gr)}.sbb.lon[data-lp="longterm"] .sbico{color:var(--am)}
.sbdiv{width:22px;height:1px;background:var(--bd);margin:4px 0;flex-shrink:0}.sbspc{flex:1}
.sbb[data-tip]::after,.sba[data-tip]::after{content:attr(data-tip);position:absolute;left:52px;top:50%;transform:translateY(-50%);background:var(--sf);border:1px solid var(--bd);border-radius:5px;padding:5px 10px;font-family:'Barlow Condensed',sans-serif;font-size:11px;letter-spacing:1.5px;text-transform:uppercase;color:var(--tx);white-space:nowrap;opacity:0;pointer-events:none;transition:opacity .15s;z-index:10}
.sbb:hover[data-tip]::after,.sba:hover[data-tip]::after{opacity:1}
#lpO{display:none;position:fixed;inset:0;z-index:584;background:rgba(0,0,0,.3)}#lpO.show{display:block}
#lp{position:fixed;top:0;left:48px;height:100%;width:600px;background:var(--sf);border-right:1px solid var(--bd);z-index:585;transform:translateX(-100vw);transition:transform .3s cubic-bezier(.16,1,.3,1);display:flex;flex-direction:column;overflow:hidden;box-shadow:var(--sh)}
#lp.open{transform:translateX(0)}

/* ══ Auth gate (login/signup/verify/reset) — full-screen, blocks everything until a session is confirmed ══ */
#authGate{position:fixed;inset:0;z-index:10000;background:var(--bg);display:flex;align-items:center;justify-content:center;padding:20px;overflow-y:auto}
#authGate.hide{display:none}
.authCard{width:380px;max-width:100%;background:var(--sf);border:1px solid var(--bd);border-radius:14px;padding:28px 26px;box-shadow:var(--sh)}
.authLogo{text-align:center;font-family:'Bebas Neue',sans-serif;font-size:24px;letter-spacing:2px;margin-bottom:4px;color:var(--ac)}
.authSub{text-align:center;font-size:11px;color:var(--dm);margin-bottom:18px}
.authTabs{display:flex;gap:4px;margin-bottom:18px;background:var(--bg);border-radius:8px;padding:3px}
.authTab{flex:1;text-align:center;padding:8px;border-radius:6px;font-size:12px;cursor:pointer;color:var(--dm);font-family:'Barlow Condensed',sans-serif;letter-spacing:.5px}
.authTab.sel{background:var(--sf2);color:var(--tx);font-weight:600}
.authScreen{display:none}.authScreen.on{display:block}
.authErr{background:rgba(255,90,90,.12);color:var(--rd);border-radius:6px;padding:8px 10px;font-size:11px;margin-bottom:12px;display:none;line-height:1.5}
.authErr.show{display:block}
.authMsg{background:rgba(77,166,255,.12);color:var(--ac);border-radius:6px;padding:8px 10px;font-size:11px;margin-bottom:12px;display:none;line-height:1.5}
.authMsg.show{display:block}
.policyBox{max-height:130px;overflow-y:auto;background:var(--bg);border:1px solid var(--bd);border-radius:8px;padding:10px;font-size:10px;line-height:1.65;color:var(--dm);margin:4px 0 12px}
.policyBox b{color:var(--tx)}
.authChk{display:flex;align-items:flex-start;gap:8px;font-size:11px;color:var(--dm);margin-bottom:14px;cursor:pointer;line-height:1.5}
.authChk input{margin-top:2px;flex-shrink:0}
.authBtn{width:100%;padding:11px;border-radius:8px;border:none;background:var(--ac);color:#0a0a0a;font-weight:700;font-size:13px;cursor:pointer;font-family:inherit;margin-top:4px}
.authBtn:disabled{opacity:.5;cursor:not-allowed}
.authLink{text-align:center;font-size:11px;color:var(--dm);margin-top:14px}
.authLink a{color:var(--ac);cursor:pointer;text-decoration:underline}
.authFld{margin-bottom:11px}
.authFld label{display:block;font-size:10px;color:var(--dm);margin-bottom:4px;font-family:'Barlow Condensed',sans-serif;letter-spacing:.5px;text-transform:uppercase}

/* ══ Watchlist / Portfolio — full-width tabs (not slide-out panels) ══ */
.addrow{padding:12px 16px;display:flex;gap:8px;flex-wrap:wrap;align-items:flex-start}
.addrow input{flex:1;min-width:110px}
.acwrap{position:relative;flex:2;min-width:180px}
#acBox{position:fixed;background:var(--sf2);border:1px solid var(--bd);border-radius:8px;max-height:220px;overflow-y:auto;z-index:9500;display:none;box-shadow:var(--sh)}
#acBox.show{display:block}
.acItem{padding:8px 11px;cursor:pointer;font-size:11px;border-bottom:1px solid var(--bd);display:flex;justify-content:space-between;gap:8px}
.acItem:last-child{border-bottom:none}
.acItem:hover,.acItem.hi{background:var(--sf)}
.acSym{font-weight:700;color:var(--tx)}
.acName{color:var(--dm);font-size:10px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-align:right}
.rmBtn{width:24px;height:24px;border:none;background:none;color:var(--dm);cursor:pointer;font-size:13px;border-radius:4px}
.rmBtn:hover{color:var(--rd);background:rgba(255,90,90,.1)}
.lphd{padding:10px 13px;border-bottom:1px solid var(--bd);background:var(--sf2);display:flex;align-items:center;gap:7px;flex-shrink:0;min-height:44px}
#obEdge{position:fixed;top:0;left:48px;right:0;height:6px;z-index:700;cursor:pointer}
#obOverlay{display:none;position:fixed;inset:0;z-index:698;background:rgba(0,0,0,.3)}
#obOverlay.show{display:block}
#obPanel{position:fixed;top:0;left:48px;right:0;height:340px;background:var(--sf);border-bottom:2px solid var(--bd);z-index:699;display:flex;flex-direction:column;transform:translateY(-100%);transition:transform .3s cubic-bezier(.16,1,.3,1);box-shadow:var(--sh)}
#obPanel.open{transform:translateY(0)}
.obHd{padding:8px 14px;border-bottom:1px solid var(--bd);background:var(--sf2);display:flex;align-items:center;gap:8px;flex-shrink:0;min-height:42px}
.obTtl{font-family:'Bebas Neue',sans-serif;font-size:15px;letter-spacing:2px;color:var(--ac)}
.lpld{width:7px;height:7px;border-radius:50%;background:var(--ac);animation:pulse 1.6s ease-in-out infinite;flex-shrink:0}
.lpttl{font-family:'Bebas Neue',sans-serif;font-size:16px;letter-spacing:2px;color:var(--ac);flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.lpib,.lpx{width:28px;height:28px;background:none;border:1px solid var(--bd);border-radius:4px;color:var(--dm);cursor:pointer;font-size:13px;display:flex;align-items:center;justify-content:center;flex-shrink:0;transition:all .2s}
.lpib:hover{border-color:var(--ac);color:var(--ac)}.lpib.on{border-color:var(--ac);color:var(--ac);background:var(--ad)}.lpx:hover{border-color:var(--rd);color:var(--rd)}
.lptm{font-family:'Barlow Condensed',sans-serif;font-size:9px;letter-spacing:1px;color:var(--dm);flex-shrink:0}
.lpSI{max-width:0;overflow:hidden;opacity:0;transition:max-width .3s,opacity .3s;flex-shrink:0}.lpSI.on{max-width:150px;opacity:1}
.lpSI input{width:135px;padding:4px 8px;background:var(--sf2);border:1px solid var(--bd);border-radius:4px;color:var(--tx);font-family:'IBM Plex Mono',monospace;font-size:11px;outline:none}
.lpSI input:focus{border-color:var(--ac)}.lpSI input::placeholder{color:var(--dm)}
.lptabs{display:flex;border-bottom:1px solid var(--bd);background:var(--sf2);flex-shrink:0}
.lptab{flex:1;padding:8px 4px;background:transparent;border:none;border-bottom:2px solid transparent;color:var(--dm);font-family:'Barlow Condensed',sans-serif;font-size:10px;letter-spacing:1px;text-transform:uppercase;cursor:pointer;transition:all .2s}
.lptab:hover{color:var(--tx)}.la-id{color:var(--cy)!important;border-bottom-color:var(--cy)!important}.la-st{color:var(--gr)!important;border-bottom-color:var(--gr)!important}.la-lt{color:var(--am)!important;border-bottom-color:var(--am)!important}
.ptsub-on{color:var(--ac)!important;border-bottom-color:var(--ac)!important}
.lpbdy{flex:1;overflow:auto;-webkit-overflow-scrolling:touch}
.lpbdy table{width:100%;border-collapse:collapse;min-width:300px}
.lpbdy thead th{padding:6px 9px;font-family:'Barlow Condensed',sans-serif;font-size:9px;letter-spacing:2px;color:var(--mu);text-align:left;text-transform:uppercase;background:var(--sf2);border-bottom:1px solid var(--bd);position:sticky;top:0;z-index:1;white-space:nowrap}
.lpbdy tbody tr{border-bottom:1px solid var(--bd);transition:background .15s}.lpbdy tbody tr:hover{background:var(--ad)}.lpbdy td{padding:6px 9px;font-size:11px;white-space:nowrap}
.lpft{padding:5px 13px;border-top:1px solid var(--bd);background:var(--sf2);font-family:'Barlow Condensed',sans-serif;font-size:9px;letter-spacing:1px;color:var(--mu);flex-shrink:0}
.sym a,.sym-lnk{color:inherit;text-decoration:none;font-weight:600}.sym a:hover,.sym-lnk:hover{color:var(--ac);text-decoration:underline}
#loader{position:fixed;inset:0;background:var(--bg);display:flex;flex-direction:column;align-items:center;justify-content:center;gap:16px;z-index:9999;transition:opacity .5s}
#loader.gn{opacity:0;pointer-events:none}
.llo{font-family:'Bebas Neue',sans-serif;font-size:38px;letter-spacing:4px;color:var(--ac)}
.lbar{width:180px;height:2px;background:var(--sf2);border-radius:2px;overflow:hidden}
.lfil{height:100%;width:0;background:var(--ac);animation:load 2s ease forwards}
.ltxt{font-family:'Barlow Condensed',sans-serif;font-size:11px;letter-spacing:3px;color:var(--dm);text-transform:uppercase}
.shell{max-width:1440px;margin:0 auto;padding:14px 22px 20px;display:grid;gap:13px}
header{display:flex;align-items:center;gap:10px;border-bottom:1px solid var(--bd);padding-bottom:11px}
.hl{display:flex;align-items:center;gap:12px;flex-shrink:0}
.lgrp{display:flex;flex-direction:column;gap:1px;flex-shrink:0}
.logo{font-family:'Bebas Neue',sans-serif;font-size:22px;letter-spacing:3px;color:var(--ac);line-height:1}
.slo{font-family:'Barlow Condensed',sans-serif;font-size:9px;letter-spacing:3px;color:var(--dm);text-transform:uppercase}
.tnav{display:flex;align-items:stretch;border:1px solid var(--bd);border-radius:7px;overflow:hidden;background:var(--sf2)}
.tbtn{padding:0 11px;height:30px;background:transparent;border:none;border-right:1px solid var(--bd);color:var(--dm);font-family:'Barlow Condensed',sans-serif;font-size:10px;letter-spacing:2px;text-transform:uppercase;cursor:pointer;display:flex;align-items:center;gap:3px;white-space:nowrap;position:relative;transition:all .2s;flex-shrink:0}
.tbtn:last-child{border-right:none}.tbtn::after{content:"";position:absolute;bottom:0;left:0;right:0;height:2px;transition:background .2s}
.tbtn:hover{color:var(--tx)}
.tbtn.active[data-tab="home"]{color:#4da6ff;background:rgba(77,166,255,.07)}.tbtn.active[data-tab="home"]::after{background:#4da6ff}
.tbtn.active[data-tab="intraday"]{color:var(--cy);background:rgba(0,212,255,.07)}.tbtn.active[data-tab="intraday"]::after{background:var(--cy)}
.tbtn.active[data-tab="shortterm"]{color:var(--gr);background:rgba(0,230,138,.07)}.tbtn.active[data-tab="shortterm"]::after{background:var(--gr)}
.tbtn.active[data-tab="longterm"]{color:var(--am);background:rgba(255,169,77,.07)}.tbtn.active[data-tab="longterm"]::after{background:var(--am)}
.hsp{flex:1}
.tcks{display:flex;align-items:center;gap:10px;flex-shrink:0}
.htck{display:flex;align-items:center;gap:4px;white-space:nowrap}
.htck-sep{width:1px;height:11px;background:var(--bd);flex-shrink:0}
.htck-n{font-family:'Barlow Condensed',sans-serif;font-size:9px;letter-spacing:1px;color:var(--mu);text-transform:uppercase}
.htck-p{font-family:'IBM Plex Mono',monospace;font-size:11px;font-weight:600;color:var(--tx)}
.htck-c{font-family:'Barlow Condensed',sans-serif;font-size:9px}.htck-c.pos{color:var(--gr)}.htck-c.neg{color:var(--rd)}
.hr{display:flex;align-items:center;gap:8px;flex-shrink:0}
.livb{display:flex;align-items:center;gap:5px;font-family:'Barlow Condensed',sans-serif;font-size:10px;letter-spacing:2px;color:var(--ac);text-transform:uppercase}
.livd{width:6px;height:6px;border-radius:50%;background:var(--ac);animation:pulse 1.6s ease-in-out infinite}
.clk{font-size:10px;color:var(--dm);letter-spacing:1px}
.icb{width:26px;height:26px;background:var(--sf2);border:1px solid var(--bd);border-radius:5px;cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:13px;color:var(--dm);transition:all .2s}.icb:hover{border-color:var(--ac);color:var(--ac)}
.mob-ham,.mob-ai{width:32px;height:32px;background:var(--sf2);border:1px solid var(--bd);border-radius:6px;cursor:pointer;display:none;align-items:center;justify-content:center;font-size:16px;color:var(--dm);flex-shrink:0}
.pbar{background:var(--sf);border:1px solid var(--bd);border-radius:6px;padding:7px 14px;display:flex;align-items:center;gap:8px;overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:none;flex-wrap:nowrap}
.pbar::-webkit-scrollbar{display:none}
.plbl{font-family:'Barlow Condensed',sans-serif;font-size:10px;letter-spacing:2px;color:var(--dm);text-transform:uppercase;white-space:nowrap;flex-shrink:0}
.mgrp{display:flex;gap:3px;flex-shrink:0}
.mbtn{padding:3px 9px;border:1px solid var(--bd);background:transparent;border-radius:4px;color:var(--dm);font-family:'Barlow Condensed',sans-serif;font-size:10px;letter-spacing:1.5px;cursor:pointer;text-transform:uppercase;white-space:nowrap;transition:all .2s;flex-shrink:0}
.mbtn:hover{border-color:var(--ac);color:var(--ac)}.mbtn.active{background:var(--ad);border-color:var(--ac);color:var(--ac)}.mbtn:disabled{opacity:.3;cursor:not-allowed}
.pinp{width:46px;padding:3px 6px;background:var(--sf2);border:1px solid var(--bd);border-radius:4px;color:var(--tx);font-family:'IBM Plex Mono',monospace;font-size:12px;text-align:center;outline:none;flex-shrink:0}
.pinp:focus{border-color:var(--ac)}.pinp:disabled{opacity:.4}
.bsp{width:1px;height:14px;background:var(--bd);flex-shrink:0}
.srng{font-size:10px;color:var(--am);letter-spacing:1px;white-space:nowrap;flex-shrink:0}
.dnav-grp{display:none;align-items:center;gap:5px;flex-shrink:0}.dnav-grp.on{display:flex}
.dnb{padding:3px 8px;border:1px solid var(--bd);background:transparent;border-radius:4px;color:var(--dm);font-family:'Barlow Condensed',sans-serif;font-size:10px;cursor:pointer;transition:all .2s;white-space:nowrap;flex-shrink:0}
.dnb:hover{border-color:var(--ac);color:var(--ac)}.dnb.tod{border-color:var(--ac);color:var(--ac);background:var(--ad)}
.datei,.timei{padding:3px 6px;background:var(--sf2);border:1px solid var(--bd);border-radius:4px;color:var(--tx);font-family:'IBM Plex Mono',monospace;font-size:10px;outline:none;flex-shrink:0}
.timei{width:72px}.timei:disabled{opacity:.4}
.ptog{display:flex;gap:3px;background:var(--sf2);border:1px solid var(--bd);border-radius:5px;padding:3px;flex-shrink:0}
.ptb{padding:3px 9px;border:none;background:transparent;border-radius:3px;color:var(--dm);font-family:'Barlow Condensed',sans-serif;font-size:10px;letter-spacing:1.5px;text-transform:uppercase;cursor:pointer;transition:all .2s}
.ptb:hover{color:var(--tx)}.ptb.active{background:var(--ac);color:#000;font-weight:700}
.mkt-tag{font-family:'Barlow Condensed',sans-serif;font-size:9px;letter-spacing:2px;color:var(--am);background:rgba(255,169,77,.1);border:1px solid var(--am);border-radius:3px;padding:2px 6px;text-transform:uppercase;flex-shrink:0}
.krow{display:grid;grid-template-columns:repeat(6,1fr);gap:10px}
.kpi{background:var(--sf);border:1px solid var(--bd);border-radius:6px;padding:11px 13px;position:relative;overflow:hidden;cursor:pointer;user-select:none;transition:border-color .2s,background .2s;display:flex;align-items:center;justify-content:space-between;gap:8px}
.kpiTxt{min-width:0}
.kring{position:relative;width:52px;height:52px;flex-shrink:0}
.kring svg{width:100%;height:100%;transform:rotate(-90deg)}
.kringBg{fill:none;stroke:var(--bd);stroke-width:3.5}
.kringFg{fill:none;stroke:var(--ka,var(--ac));stroke-width:3.5;stroke-linecap:round;transition:stroke-dashoffset .6s cubic-bezier(.2,.7,.3,1)}
.kringPct{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;font-family:'Bebas Neue',sans-serif;font-size:14px;color:var(--ka,var(--ac));letter-spacing:.5px}
.kpi::before{content:"";position:absolute;top:0;left:0;width:3px;height:100%;background:var(--ka,var(--ac));transition:width .2s}
.kpi:hover{border-color:var(--ka,var(--ac));background:var(--sf2)}.kpi:hover::before{width:5px}
.kpi.af{border-color:var(--ka,var(--ac));box-shadow:0 0 0 1px var(--ka,var(--ac))}
.klbl{font-family:'Barlow Condensed',sans-serif;font-size:9px;letter-spacing:2px;color:var(--dm);text-transform:uppercase;margin-bottom:4px}
.kval{font-family:'Bebas Neue',sans-serif;font-size:25px;color:var(--ka,var(--ac));line-height:1}
.ksub{font-size:9px;color:var(--dm);margin-top:3px}
.mgrid{display:grid;grid-template-columns:1fr 1fr 300px;gap:13px}
.pan{background:var(--sf);border:1px solid var(--bd);border-radius:6px;overflow:hidden}
.phd{padding:8px 13px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid var(--bd)}
.pttl{font-family:'Barlow Condensed',sans-serif;font-size:10px;letter-spacing:2.5px;color:var(--dm);text-transform:uppercase}
.bdg{font-size:9px;padding:2px 7px;border-radius:16px;font-family:'Barlow Condensed',sans-serif;letter-spacing:1px}
.bg{background:var(--gdm);color:var(--gr)}.br{background:var(--rdm);color:var(--rd)}.ba{background:rgba(255,169,77,.10);color:var(--am)}.bb{background:rgba(77,166,255,.10);color:var(--bl)}.bc{background:rgba(0,212,255,.10);color:var(--cy)}
table{width:100%;border-collapse:collapse}
thead th{padding:6px 10px;font-family:'Barlow Condensed',sans-serif;font-size:9px;letter-spacing:1.5px;color:var(--mu);text-align:left;text-transform:uppercase;background:var(--sf2);border-bottom:1px solid var(--bd);white-space:nowrap}
tbody tr{border-bottom:1px solid var(--bd);transition:background .15s}tbody tr:last-child{border-bottom:none}tbody tr:hover{background:var(--ad)}
td{padding:6px 10px;font-size:11px}
.sym{font-weight:600;color:var(--tx)}.rp{color:var(--gr);font-weight:600}.rn{color:var(--rd);font-weight:600}
.emr td{text-align:center;color:var(--dm);padding:22px;font-family:'Barlow Condensed',sans-serif;letter-spacing:2px;font-size:10px}
.pill{display:inline-block;padding:2px 6px;border-radius:16px;font-family:'Barlow Condensed',sans-serif;font-size:9px;letter-spacing:1px}
.pbuy{background:rgba(0,230,138,.12);color:var(--gr)}.psell{background:rgba(255,63,91,.12);color:var(--rd)}.prun{background:rgba(255,169,77,.12);color:var(--am)}.psuc{background:var(--gdm);color:var(--gr)}.pfal{background:var(--rdm);color:var(--rd)}.pcls{background:rgba(180,142,255,.12);color:var(--pu)}
.cbdg{display:inline-block;font-size:8px;letter-spacing:1px;background:rgba(77,166,255,.15);color:var(--bl);border-radius:3px;padding:1px 4px;margin-left:3px;font-family:'Barlow Condensed',sans-serif;text-transform:uppercase;vertical-align:middle}
.symacts{display:inline-flex;align-items:center;gap:2px;vertical-align:middle;margin-left:3px}
.symacts span{display:inline-flex;align-items:center;justify-content:center;width:13px;height:13px;font-size:9px;cursor:pointer;background:transparent;border:none;color:var(--mu);transition:color .12s;flex-shrink:0;line-height:1;vertical-align:middle}
.symacts .nwi:hover,.symacts .obi:hover,.symacts .noi:hover{color:var(--tx)}
.wli{display:inline-block;vertical-align:middle;margin-right:3px;font-size:11px;cursor:pointer;color:var(--mu);transition:color .12s;line-height:1}
.wli:hover{color:var(--am)}
.wli.on{color:var(--am)}
.nwi:hover,.nwi.on{opacity:1}
#nwBox{position:fixed;background:var(--sf);border:1px solid var(--bd);border-radius:8px;padding:10px 12px;width:280px;max-width:80vw;max-height:260px;overflow-y:auto;font-family:'Barlow Condensed',sans-serif;font-size:11px;line-height:1.6;color:var(--tx);text-align:left;z-index:9999;box-shadow:var(--sh);display:none}
#nwBox.show{display:block}
#nwBox a{color:var(--ac);text-decoration:underline}
#nwBox hr.nwdiv{border:none;border-top:1px solid var(--bd);margin:7px 0}
#noBox{position:fixed;background:var(--sf);border:1px solid var(--bd);border-radius:8px;padding:10px 12px;width:240px;max-width:80vw;max-height:200px;overflow-y:auto;font-family:'Barlow Condensed',sans-serif;font-size:11px;line-height:1.6;color:var(--tx);text-align:left;z-index:9999;box-shadow:var(--sh);display:none;white-space:pre-wrap}
#noBox.show{display:block}
.sym.symfl{position:relative;padding-left:15px}
.sym.symfl::before{content:"";position:absolute;left:3px;top:3px;bottom:3px;width:3px;border-radius:2px;background:var(--am)}
.ccol{display:flex;flex-direction:column;gap:12px}
.dwrap{position:relative;width:110px;height:110px;margin:10px auto 4px}
.dwrap svg{width:100%;height:100%;transform:rotate(-90deg)}
.dctr{position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center}
.dpct{font-family:'Bebas Neue',sans-serif;font-size:26px;color:var(--gr);line-height:1}
.dlbl{font-size:9px;letter-spacing:2px;color:var(--dm);text-transform:uppercase}
.legi{display:flex;flex-direction:column;gap:5px;padding:0 12px 8px}
.li{display:flex;align-items:center;gap:6px}.ld{width:7px;height:7px;border-radius:2px;flex-shrink:0}.lt{flex:1;font-size:10px;color:var(--dm)}.lv{font-weight:600;font-size:10px}
.bsec{padding:8px 12px}
.brow{display:flex;align-items:center;gap:6px;margin-bottom:6px}
.blbl{width:64px;font-size:9px;color:var(--dm);text-align:right;font-family:'Barlow Condensed',sans-serif;letter-spacing:1px;flex-shrink:0}
.btrk{flex:1;height:4px;background:var(--sf2);border-radius:3px;overflow:hidden}
.bfil{height:100%;border-radius:3px;width:0;transition:width 1s cubic-bezier(.16,1,.3,1)}
.bnum{width:22px;text-align:right;font-size:10px;font-weight:600;flex-shrink:0}
#dr,#idr,#ltdr{background:var(--sf);border:1px solid var(--bd);border-radius:6px;overflow:hidden;max-height:0;opacity:0;transition:max-height .4s cubic-bezier(.16,1,.3,1),opacity .3s}
#dr.open,#idr.open,#ltdr.open{max-height:600px;opacity:1}
.drhd{padding:8px 14px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid var(--bd);background:var(--sf2)}
.drttl{font-family:'Bebas Neue',sans-serif;font-size:15px;letter-spacing:2px;color:var(--drc,var(--ac))}.drcnt{font-family:'Barlow Condensed',sans-serif;font-size:10px;letter-spacing:2px;color:var(--dm)}
.drx{background:none;border:1px solid var(--bd);border-radius:4px;color:var(--dm);width:26px;height:26px;font-size:13px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all .2s}.drx:hover{border-color:var(--rd);color:var(--rd)}
.drsc{margin:6px 13px;width:calc(100% - 26px);padding:6px 10px;background:var(--sf2);border:1px solid var(--bd);border-radius:4px;color:var(--tx);font-family:'IBM Plex Mono',monospace;font-size:11px;outline:none;display:block}
.drsc:focus{border-color:var(--ac)}.drsc::placeholder{color:var(--dm)}
.drbdy th[data-col]{cursor:pointer;user-select:none;white-space:nowrap}
.drbdy th[data-col]:hover{color:var(--tx)}
.drbdy th[data-col]::after{content:"";display:inline-block;width:0;height:0;margin-left:4px;vertical-align:middle;border-left:3px solid transparent;border-right:3px solid transparent;opacity:.35}
.drbdy th[data-col].srt-asc::after{border-bottom:4px solid var(--ac);opacity:1}
.drbdy th[data-col].srt-desc::after{border-top:4px solid var(--ac);opacity:1}
.drbdy{overflow:auto;max-height:440px;-webkit-overflow-scrolling:touch}
.drbdy thead th{padding:5px 9px;font-family:'Barlow Condensed',sans-serif;font-size:9px;letter-spacing:1.5px;color:var(--mu);text-align:left;text-transform:uppercase;background:var(--sf2);border-bottom:1px solid var(--bd);position:sticky;top:0;white-space:nowrap}
/* Frozen headers, corrected: the first attempt scoped sticky only by container id, but a table's
   own overflow-x:auto wrapper (needed for horizontal scroll on narrow screens) becomes the nearest
   scrolling ancestor for position:sticky — and since that wrapper's height was unconstrained, it
   never scrolled internally itself (the whole page scrolled instead, carrying the wrapper along as
   one rigid block), so nothing ever actually stuck. Giving the wrapper a real bounded height +
   overflow-y — the same pattern .lpbdy/.drbdy above already use successfully — makes it a genuine
   scroll container, which is what sticky actually needs. */
.scrollTbl{overflow-x:auto;overflow-y:auto;max-height:min(620px,65vh)}
/* The outer table's own header row is the only thing that freezes. NOTE: the selector must go
   through tr — `thead > th` matches nothing at all, since th is a child of tr, not of thead. */
.scrollTbl>table>thead>tr>th{position:sticky;top:0;z-index:2;background:var(--sf2)}
/* Both header rows freeze. The dropdown's header stacks directly BELOW the outer one rather than
   fighting it for top:0 — the offset is measured at runtime into --outerThH (see syncStickyOffsets_)
   instead of being a guessed pixel value, so it stays correct across font sizes and zoom levels.
   Lower z-index than the outer header so it slides under it, never over. */
.scrollTbl td table thead th{position:sticky;top:var(--outerThH,28px);z-index:1;background:var(--sf2)}
/* Order Book drawer — same freeze. Its wrapper already had a bounded height (flex:1 inside the
   drawer's flex column), so it only needed overflow-y plus this rule. */
.obScroll>table>thead>tr>th{position:sticky;top:0;z-index:2;background:var(--sf2)}
.drbdy tbody tr{border-bottom:1px solid var(--bd)}.drbdy td{padding:5px 9px;font-size:11px;white-space:nowrap}
.php{background:var(--sf);border:1px solid var(--bd);border-radius:6px;padding:40px 24px;text-align:center}
.phico{font-size:44px;margin-bottom:12px;opacity:.18}.phttl{font-family:'Bebas Neue',sans-serif;font-size:22px;letter-spacing:3px;margin-bottom:7px}
.phsub{font-family:'Barlow Condensed',sans-serif;font-size:12px;letter-spacing:2px;color:var(--dm);text-transform:uppercase;margin-bottom:16px}
.phch{display:flex;justify-content:center;gap:8px;flex-wrap:wrap;margin-bottom:14px}
.phc{background:var(--sf2);border:1px solid var(--bd);border-radius:4px;padding:4px 10px;font-family:'Barlow Condensed',sans-serif;font-size:11px;letter-spacing:1px;color:var(--dm);text-transform:uppercase;cursor:pointer}
.fab{position:fixed;bottom:20px;right:18px;width:40px;height:40px;border-radius:50%;background:var(--ac);border:none;cursor:pointer;font-size:16px;display:flex;align-items:center;justify-content:center;z-index:700;box-shadow:0 3px 12px rgba(0,0,0,.4);color:#000;transition:transform .2s}.fab:hover{transform:scale(1.1)}
.mover{display:none;position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:2000;align-items:center;justify-content:center;padding:16px}.mover.open{display:flex}
.modal{background:var(--sf);border:1px solid var(--bd);border-radius:10px;width:100%;max-width:420px;max-height:88vh;overflow:hidden;box-shadow:var(--sh);display:flex;flex-direction:column;transform:scale(.96);opacity:0;transition:transform .18s cubic-bezier(.2,.7,.3,1),opacity .18s}
.mover.open .modal{transform:scale(1);opacity:1}
.mhd{padding:12px 16px;border-bottom:1px solid var(--bd);display:flex;align-items:center;justify-content:space-between;background:var(--sf2);flex-shrink:0}
.mttl{font-family:'Bebas Neue',sans-serif;font-size:18px;letter-spacing:2px;color:var(--ac)}
.mx{background:none;border:1px solid var(--bd);border-radius:4px;color:var(--dm);width:26px;height:26px;cursor:pointer;font-size:13px;display:flex;align-items:center;justify-content:center;transition:all .2s}.mx:hover{border-color:var(--rd);color:var(--rd)}
.mbdy{padding:13px 16px;overflow-y:auto;flex:1;min-height:0}
.setSec{border:1px solid var(--bd);border-radius:8px;margin-bottom:10px;overflow:hidden}
.setSec summary{padding:11px 14px;cursor:pointer;list-style:none;display:flex;align-items:center;justify-content:space-between;background:var(--sf2);font-family:'Barlow Condensed',sans-serif;font-size:12px;letter-spacing:.5px;text-transform:uppercase;color:var(--tx);font-weight:600}
.setSec summary::-webkit-details-marker{display:none}
.setSec summary::after{content:"\25BE";font-size:11px;color:var(--dm);transition:transform .2s;margin-left:8px}
.setSec[open] summary::after{transform:rotate(180deg)}
.setSec .setBody{padding:14px}
.mlbl{font-family:'Barlow Condensed',sans-serif;font-size:10px;letter-spacing:2.5px;color:var(--dm);text-transform:uppercase;margin-bottom:6px;display:block}
.mtyps{display:flex;gap:6px;margin-bottom:10px;flex-wrap:wrap}
.tybtn{padding:5px 10px;border:1px solid var(--bd);border-radius:4px;background:transparent;color:var(--dm);font-family:'Barlow Condensed',sans-serif;font-size:10px;letter-spacing:1px;cursor:pointer;text-transform:uppercase;transition:all .2s}.tybtn.sel{border-color:var(--ac);color:var(--ac);background:var(--ad)}
.minp{width:100%;padding:7px 9px;background:var(--sf2);border:1px solid var(--bd);border-radius:5px;color:var(--tx);font-family:'IBM Plex Mono',monospace;font-size:12px;outline:none;margin-bottom:8px;resize:vertical}
.minp:focus{border-color:var(--ac)}.minp::placeholder{color:var(--dm)}
.mft{display:flex;justify-content:flex-end;gap:8px;margin-top:4px}
.bcx{padding:6px 12px;background:transparent;border:1px solid var(--bd);border-radius:5px;color:var(--dm);font-family:'Barlow Condensed',sans-serif;font-size:11px;letter-spacing:1px;cursor:pointer;text-transform:uppercase}
.bok{padding:6px 14px;background:var(--ac);border:1px solid var(--ac);border-radius:5px;color:#000;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;letter-spacing:1px;cursor:pointer;text-transform:uppercase}.bok:hover{opacity:.85}
#toast{position:fixed;bottom:68px;left:50%;transform:translateX(-50%);background:var(--sf);border:1px solid var(--bd);border-radius:6px;padding:8px 16px;font-family:'Barlow Condensed',sans-serif;font-size:12px;letter-spacing:2px;color:var(--tx);opacity:0;transition:opacity .3s;pointer-events:none;z-index:3000;box-shadow:var(--sh);white-space:nowrap}
#toast.show{opacity:1}
footer{border-top:1px solid var(--bd);padding-top:9px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:6px}
.fnote{font-size:10px;color:var(--mu)}.favg{font-size:11px;color:var(--dm)}.favg .v{font-weight:600}.favg .pos{color:var(--gr)}.favg .neg{color:var(--rd)}.fupd{font-size:11px;color:var(--dm)}.fupd span{color:var(--ac)}
@keyframes fadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
@keyframes pulse{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.4;transform:scale(.7)}}
@keyframes load{to{width:100%}}
@media(max-width:1024px){
  #sbT,#aiT{display:none!important}.mob-ham,.mob-ai{display:flex!important}
  .tcks{display:none!important}
  .krow{grid-template-columns:repeat(3,1fr)!important;gap:7px}
  .kpi{padding:9px 10px}.kval{font-size:22px}
  .mgrid{grid-template-columns:1fr!important;gap:9px}
  .ccol{flex-direction:column!important}
  .shell{padding:10px 12px;gap:10px}
  #lp{left:0!important;width:min(600px,100vw)!important}
  #aip{width:min(360px,100vw)!important}
  #aipResize{display:none}
  .tnav{overflow-x:auto;scrollbar-width:none;-webkit-overflow-scrolling:touch}.tnav::-webkit-scrollbar{display:none}
  .lgrp{display:none}
  header{gap:7px}
  .mgrid .pan table,.drbdy table,.lpbdy table{display:block;overflow-x:auto;-webkit-overflow-scrolling:touch}
}
@media(max-width:540px){.krow{grid-template-columns:repeat(2,1fr)!important}.mbtn{font-size:9px;padding:3px 6px}}

/* v2 additions */
.chgTbl{width:100%;border-collapse:collapse;margin:8px 0 10px}
.chgTbl th{font-size:9px;color:var(--dm);text-align:center;font-weight:600;padding:3px 4px;text-transform:uppercase;letter-spacing:.6px;border-bottom:1px solid var(--bd)}
.chgTbl td{font-size:10px;color:var(--dm);padding:3px 4px;border-bottom:1px solid var(--bd);white-space:normal}
.chgTbl td:first-child{width:46%}
.chgIn{height:28px!important;margin:0!important;padding:2px 6px!important;font-size:11px!important;text-align:right}
.setHint{font-size:10px;color:var(--dm);line-height:1.5;margin-bottom:10px}
.setWarn{font-size:10px;color:var(--am);background:rgba(255,180,60,.1);border-radius:6px;padding:8px 10px;line-height:1.5;margin-bottom:10px}
.setDiv{border-top:1px solid var(--bd);margin:16px 0 12px}
.mktStrip{padding:7px 14px;font-size:11px;color:var(--dm);border:1px solid var(--bd);border-radius:6px;background:var(--sf);margin:4px 0}
.mktStrip:empty{display:none}
.catSel{width:auto!important;margin:0 0 0 8px!important;height:28px;padding:3px 6px!important;display:inline-block!important}
.gd{font-size:12px;line-height:1.65;color:var(--tx)}
.gd,.gd *{-webkit-user-select:text;user-select:text}
.gd h4{font-family:'Barlow Condensed',sans-serif;font-size:13px;letter-spacing:2px;text-transform:uppercase;color:var(--ac);margin:20px 0 6px;border-bottom:1px solid var(--bd);padding-bottom:4px}
.gd ul,.gd ol{margin:6px 0 8px 20px;padding:0}.gd li{margin:4px 0}
.gd code{background:var(--sf2);border:1px solid var(--bd);border-radius:4px;padding:1px 5px;font-size:11px;word-break:break-all}
.gdLead{color:var(--dm);margin-bottom:6px}
.gdBox{background:var(--sf2);border:1px solid var(--bd);border-radius:8px;padding:10px 12px;margin:10px 0;font-size:11px}
.gdOk{border-color:var(--gr);background:rgba(0,230,138,.07)}
.gdNote{font-size:11px;color:var(--dm);margin:6px 0 4px}
.gdSvg{width:100%;max-width:640px;height:auto;display:block;margin:10px auto;color:var(--tx)}
.gdSvg text{font-family:'IBM Plex Mono',monospace}
.gdTbl{width:100%;border-collapse:collapse;margin:8px 0;font-size:11px}
.gdTbl th{text-align:left;padding:5px 8px;background:var(--sf2);border-bottom:1px solid var(--bd);font-size:10px;letter-spacing:1px;text-transform:uppercase;color:var(--dm)}
.gdTbl td{padding:5px 8px;border-bottom:1px solid var(--bd);vertical-align:top}
.gdCode{position:relative;margin:8px 0}.gdCode pre{background:#0b0f14;color:#cfe3d6;border:1px solid var(--bd);border-radius:8px;padding:12px;font-size:10.5px;line-height:1.5;overflow:auto;max-height:340px;margin:0}
.gdCopy{position:absolute;top:6px;right:8px;z-index:2}

/* v3 layout: one compact toolbar + tab strip instead of five stacked bars */
.ptTb{display:flex;align-items:center;gap:8px;flex-wrap:wrap;padding:5px 10px;background:var(--sf);border:1px solid var(--bd);border-radius:6px;position:sticky;top:0;z-index:7}
.ptModeBar{display:flex;align-items:center;gap:6px}
.ptLbl{font-family:'Barlow Condensed',sans-serif;font-size:9px;letter-spacing:1.5px;color:var(--dm);text-transform:uppercase}
.ptModeBanner{font-family:'Barlow Condensed',sans-serif;font-size:10px;letter-spacing:1px;color:var(--rd);text-transform:uppercase}
.ptTb .bcx,.ptTb .bok{padding:3px 10px;font-size:10px;line-height:1.5}
.ptTbBtn{white-space:nowrap}
.ptTbStrip{margin:0!important;padding:0!important;border:none!important;background:none!important;font-size:10.5px;color:var(--dm);line-height:1.4}
.ptTbLive{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.liveSt{font-size:10.5px;color:var(--dm)}
.ptSp{flex:1}
.ptIco{width:28px;height:24px;padding:0!important;display:inline-flex;align-items:center;justify-content:center}
.ptstrip{align-items:center;flex-wrap:wrap;margin-top:6px;border:1px solid var(--bd);border-bottom:none;border-radius:6px 6px 0 0}
.ptstrip .lptab{flex:0 0 auto;padding:7px 14px}
.ptstrip .lptab .bdg{margin-left:4px}
.ptstrip .drsc{margin:3px 4px;width:170px;height:24px;padding:2px 8px;display:inline-block}
.ptstrip .catSel{width:auto;margin-right:8px!important}
.ptBody{border-radius:0 0 6px 6px;border-top:none}
.ptSum{padding:5px 12px;font-size:10.5px;color:var(--dm);border-bottom:1px solid var(--bd);line-height:1.6}
.ptSum:empty{display:none}
#cnt-portfolio .scrollTbl,#cnt-watchlist .scrollTbl{max-height:calc(100dvh - 205px);min-height:200px}
.wlBar{flex-wrap:wrap;gap:8px;justify-content:flex-start;align-items:center}
.wlAdd{display:flex;align-items:center;gap:6px;flex:1;min-width:260px}
.wlAdd .acwrap{flex:2;min-width:150px}
.wlAdd .minp{height:28px;margin:0;padding:3px 8px}
.wlNoteIn{flex:1.2;min-width:100px}
.wlSt{font-size:10.5px;white-space:nowrap}
.wlSearch{margin:0!important;width:180px!important;height:28px;padding:4px 9px!important;display:inline-block!important}
.guideModal{max-width:780px;width:96vw}
.guideBdy{max-height:calc(88vh - 56px)}
@media(max-width:700px){.ptstrip .drsc{width:120px}.ptTb{gap:6px}}
.gdBtns{position:absolute;top:6px;right:8px;z-index:2;display:flex;gap:6px}
.gdFile{font-size:11px;color:var(--dm);margin:12px 0 2px}
.gdCode pre{padding-top:34px}
</style>
</head>
<body>
<div id="authGate">
  <div class="authCard">
    <div class="authLogo">&#10024; SIGNAL DASHBOARD</div>
    <div class="authSub">Sign in to track your signals, watchlist &amp; portfolio</div>

    <div class="authTabs" id="authTabs">
      <div class="authTab sel" onclick="showAuthScreen('login')">Log In</div>
      <div class="authTab" onclick="showAuthScreen('signup')">Sign Up</div>
    </div>

    <!-- LOGIN -->
    <div class="authScreen on" id="authLogin">
      <div class="authErr" id="loginErr"></div>
      <div class="authFld"><label>Email</label><input class="minp" id="loginEmail" type="email" style="height:36px" placeholder="you@email.com" onkeydown="if(event.key==='Enter')doLogin();"/></div>
      <div class="authFld"><label>Password</label><input class="minp" id="loginPass" type="password" style="height:36px" placeholder="••••••••" onkeydown="if(event.key==='Enter')doLogin();"/></div>
      <button class="authBtn" id="loginBtn" onclick="doLogin()">Log In</button>
      <div class="authLink"><a onclick="showAuthScreen('forgot')">Forgot password?</a></div>
    </div>

    <!-- SIGNUP -->
    <div class="authScreen" id="authSignup">
      <div class="authErr" id="signupErr"></div>
      <div class="authFld"><label>Name</label><input class="minp" id="suName" type="text" style="height:36px" placeholder="Your name"/></div>
      <div class="authFld"><label>Email</label><input class="minp" id="suEmail" type="email" style="height:36px" placeholder="you@email.com"/></div>
      <div class="authFld"><label>Password</label><input class="minp" id="suPass" type="password" style="height:36px" placeholder="At least 8 characters"/></div>
      <div class="authFld"><label>Confirm Password</label><input class="minp" id="suPass2" type="password" style="height:36px" placeholder="Repeat password"/></div>
      <label style="display:block;font-size:10px;color:var(--dm);margin-bottom:4px;font-family:'Barlow Condensed',sans-serif;letter-spacing:.5px;text-transform:uppercase">Privacy Policy</label>
      <div class="policyBox">
        <b>What we collect:</b> your email, display name, a securely hashed password (never stored in plain text), and any watchlist/portfolio entries you create.<br/><br/>
        <b>How we use it:</b> to create and secure your account and show your own data back to you. We don't use it for advertising and don't sell or share it with third parties.<br/><br/>
        <b>Where it lives:</b> a Google Sheet in the operator's own Google account, processed via Google Apps Script — not an independent third-party server. AI questions you ask are sent to Groq's API to generate a response.<br/><br/>
        <b>Who can access it:</b> your password is hashed and never visible to anyone, including the operator. The operator does have underlying technical access to the sheet itself, since this runs on their own Google account.<br/><br/>
        <b>Deletion:</b> you can delete your account and all associated data anytime from your Profile.<br/><br/>
        <b>Not financial advice:</b> signals and AI commentary are informational only — always consult a qualified financial advisor.
      </div>
      <label class="authChk"><input type="checkbox" id="suAgree"/> I have read and agree to the Privacy Policy above</label>
      <button class="authBtn" id="signupBtn" onclick="doSignup()">Create Account</button>
    </div>

    <!-- VERIFY EMAIL -->
    <div class="authScreen" id="authVerify">
      <div class="authErr" id="verifyErr"></div>
      <div class="authMsg show" id="verifyMsg">We've sent a 6-digit code to your email. Enter it below to finish creating your account.</div>
      <div class="authFld"><label>Verification Code</label><input class="minp" id="verifyOtp" type="text" inputmode="numeric" maxlength="6" style="height:36px;text-align:center;font-size:18px;letter-spacing:4px" placeholder="------"/></div>
      <button class="authBtn" id="verifyBtn" onclick="doVerify()">Verify &amp; Continue</button>
      <div class="authLink"><a onclick="doResend()">Resend code</a></div>
    </div>

    <!-- TWO-STEP VERIFICATION (v2) -->
    <div class="authScreen" id="authMfa">
      <div class="authErr" id="mfaErr"></div>
      <div class="authMsg show">Two-step verification: enter the 6-digit code from your authenticator app. Lost your phone? Enter one of your recovery codes instead.</div>
      <div class="authFld"><label>Verification code</label><input class="minp" id="mfaCode" type="text" autocomplete="one-time-code" maxlength="12" style="height:36px;text-align:center;font-size:18px;letter-spacing:3px" placeholder="------" onkeydown="if(event.key==='Enter')doMfaVerify();"/></div>
      <button class="authBtn" id="mfaBtn" onclick="doMfaVerify()">Verify &amp; Log In</button>
      <div class="authLink"><a onclick="mfaBackToLogin()">Back to log in</a></div>
    </div>

    <!-- FORGOT PASSWORD -->
    <div class="authScreen" id="authForgot">
      <div class="authErr" id="forgotErr"></div>
      <div class="authMsg" id="forgotMsg"></div>
      <div class="authFld"><label>Email</label><input class="minp" id="forgotEmail" type="email" style="height:36px" placeholder="you@email.com"/></div>
      <button class="authBtn" id="forgotBtn" onclick="doForgot()">Send Reset Code</button>
      <div class="authLink"><a onclick="showAuthScreen('login')">Back to log in</a></div>
    </div>

    <!-- RESET PASSWORD -->
    <div class="authScreen" id="authReset">
      <div class="authErr" id="resetErr"></div>
      <div class="authFld"><label>Reset Code</label><input class="minp" id="resetOtp" type="text" inputmode="numeric" maxlength="6" style="height:36px;text-align:center;font-size:18px;letter-spacing:4px" placeholder="------"/></div>
      <div class="authFld"><label>New Password</label><input class="minp" id="resetPass" type="password" style="height:36px" placeholder="At least 8 characters"/></div>
      <div class="authFld"><label>Confirm New Password</label><input class="minp" id="resetPass2" type="password" style="height:36px" placeholder="Repeat new password"/></div>
      <button class="authBtn" id="resetBtn" onclick="doReset()">Reset Password</button>
    </div>
  </div>
</div>
<div id="nwBox"></div>
<div id="noBox"></div>
<div id="acBox"></div>
<div id="obEdge"></div>
<div id="obOverlay" onclick="closeOB()"></div>
<div id="obPanel">
  <div class="obHd">
    <span class="obTtl" id="obTtl">&#9654; Order Book</span>
    <span class="bdg ba" id="obCount" style="margin-left:8px">0</span>
    <span id="obModeTabs" style="display:none;gap:4px;margin-left:10px"><button class="bcx" id="obModeMock" style="padding:2px 9px;font-size:10px" onclick="setObMode('mock')">Mock</button><button class="bcx" id="obModeLive" style="padding:2px 9px;font-size:10px" onclick="setObMode('live')">Live</button></span>
    <div style="display:flex;gap:6px;align-items:center;margin-left:auto">
      <input class="minp" id="obSearch" placeholder="Symbol..." style="height:26px;max-width:100px;font-size:10px" oninput="filterOrders()"/>
      <select class="minp" id="obFilterStatus" style="height:26px;font-size:10px;max-width:90px" onchange="filterOrders()">
        <option value="">All</option><option value="PENDING">Pending</option><option value="EXECUTED">Executed</option><option value="FILLED">Filled</option><option value="CANCELLED">Cancelled</option>
      </select>
      <select class="minp" id="obFilterCat" style="height:26px;font-size:10px;max-width:100px" onchange="filterOrders()">
        <option value="">All Tabs</option><option value="intraday">Intraday</option><option value="shortterm">Short-term</option><option value="longterm">Long-term</option>
      </select>
      <select class="minp" id="obFilterStatusLive" style="display:none;height:26px;font-size:10px;max-width:90px" onchange="filterOrders()"><option value="">All</option><option value="OPEN">Open</option><option value="COMPLETE">Complete</option><option value="REJECTED">Rejected</option><option value="CANCELLED">Cancelled</option></select>
      <select class="minp" id="obFilterProdLive" style="display:none;height:26px;font-size:10px;max-width:100px" onchange="filterOrders()"><option value="">All products</option><option value="MIS">Intraday</option><option value="CNC">Delivery</option></select>
      <button class="bok" style="height:26px;padding:0 10px;font-size:10px" onclick="loadOrders()">&#8635;</button>
      <button class="lpx" onclick="closeOB()">&#10005;</button>
    </div>
  </div>
  <div class="obScroll" style="overflow-x:auto;overflow-y:auto;flex:1">
    <table><thead id="obHeadMock"><tr><th>#</th><th>Symbol</th><th>Category</th><th>Type</th><th>Qty</th><th>Price</th><th>Amount</th><th>Target</th><th>SL</th><th>Status</th><th>Date</th><th>Notes</th><th></th></tr></thead><thead id="obHeadLive" style="display:none"><tr><th>#</th><th>Symbol</th><th>Product</th><th>Type</th><th>Qty</th><th>Price</th><th>Amount</th><th>Status</th><th>Time</th><th></th></tr></thead>
      <tbody id="obTblBdy"><tr class="emr"><td colspan="13">Move cursor to top edge to open. &#9654; icon on any symbol to place an order.</td></tr></tbody>
    </table>
  </div>
</div>
<div id="sbT"></div><div id="sbO"></div>
<div id="sb">
  <button class="sbb" data-lp="intraday"  data-tip="Intraday Live"   onclick="openLP('intraday')"><span class="sbico">&#9889;</span></button>
  <button class="sbb" data-lp="shortterm" data-tip="Short-term Live" onclick="openLP('shortterm')"><span class="sbico">&#8645;</span></button>
  <button class="sbb" data-lp="longterm"  data-tip="Long-term Live"  onclick="openLP('longterm')"><span class="sbico">&#9651;</span></button>
  <div class="sbdiv"></div><div class="sbspc"></div>
  <div class="refBox" id="refBox" title="Share your referral link" onclick="openShare()">
    <span class="sbico" style="font-size:14px">&#128279;</span>
    <span class="refTxt" id="refTxt">Refer</span>
  </div>
  <button class="sbb" data-tip="Settings" onclick="openProfile()"><span class="sbico">&#9881;</span></button>
  <button class="sbb" id="trafficBtn" data-tip="Traffic Stats" onclick="openTraffic()" style="display:none"><span class="sbico">&#128202;</span></button>
  <button class="sba" data-tip="Toggle Theme" onclick="toggleTheme()"><span class="sbico" id="sbTI">&#9681;</span></button>
  <button class="sba" data-tip="Feedback"     onclick="openFB()"><span class="sbico">&#9993;</span></button>
</div>

<div id="aiT"></div><div id="aiO"></div>
<div id="aip">
  <div id="aipResize" title="Drag to resize"></div>
  <div class="aihd">
    <svg width="18" height="18" viewBox="0 0 24 24" fill="var(--ac)" stroke="none"><path d="M12 2l1.8 5.6L19 9.5l-5.2 1.9L12 17l-1.8-5.6L5 9.5l5.2-1.9L12 2z"/><path d="M19 13l.9 2.8L23 17l-3.1 1.2L19 21l-.9-2.8L15 17l3.1-1.2L19 13z"/></svg>
    <span class="aittl">Signal AI</span><span class="aisub">Groq</span>
    <button class="lpib" id="aiPinBtn" onclick="toggleAIPin()" title="Pin panel">&#128204;</button>
    <button class="lpx" onclick="closeAI()">&#10005;</button>
  </div>
  <div class="aibd" id="aiBd">
    <div class="ai-welcome">
      <span class="ai-title">&#10024; Signal AI</span>
      Ask about your signals, market news, stock analysis, or any trading question.
      <div class="ai-chips">
        <span class="ai-chip" onclick="askAI('Summarise today intraday session')">Today Session</span>
        <span class="ai-chip" onclick="askAI('What are the top performing signals?')">Top Signals</span>
        <span class="ai-chip" onclick="askAI('What is the current win rate?')">Win Rate</span>
        <span class="ai-chip" onclick="askAI('Latest NSE market news today')">Market News</span>
      </div>
      <div style="margin-top:10px;font-size:9px;color:var(--gr)">&#10003; Groq LLaMA 3.3 70B &middot; Web Search Enabled</div>
    </div>
  </div>
  <div class="ai-inprow">
    <textarea class="ai-inp" id="aiIn" placeholder="Ask anything about your signals..." rows="1"
      oninput="this.style.height='auto';this.style.height=Math.min(this.scrollHeight,80)+'px';"
      onkeydown="if(event.key==='Enter'&&!event.shiftKey){event.preventDefault();sendAI();}"></textarea>
    <button class="ai-send" id="aiSend" onclick="sendAI()">&#8593;</button>
  </div>
</div>

<div id="lpO"></div>
<div id="lp">
  <div class="lphd">
    <div class="lpld"></div>
    <span class="lpttl" id="lpTtl">Live Signals</span>
    <div class="lpSI" id="lpSI"><input type="text" id="lpSIn" placeholder="Search..." oninput="filterLP()"/></div>
    <button class="lpib" id="lpSW" onclick="toggleLPSearch()">&#9906;</button>
    <button class="lpib" id="lpPinBtn" onclick="toggleLPPin()" title="Pin panel">&#128204;</button>
    <span class="lptm" id="lpTm">15s</span>
    <button class="lpx" onclick="closeLP()">&#10005;</button>
  </div>
  <div class="lptabs">
    <button class="lptab la-id" data-lt="intraday"  onclick="switchLP('intraday')">&#9889; Intraday</button>
    <button class="lptab"       data-lt="shortterm" onclick="switchLP('shortterm')">&#8645; Short-term</button>
    <button class="lptab"       data-lt="longterm"  onclick="switchLP('longterm')">&#9651; Long-term</button>
  </div>
  <div class="lpbdy" id="lpBdy"><div style="text-align:center;padding:30px;color:var(--dm);font-family:'Barlow Condensed',sans-serif;letter-spacing:2px;font-size:10px">LOADING...</div></div>
  <div class="lpft">Hover symbol for name &middot; Click to open Screener.in &middot; Auto-refresh 15s</div>
</div>



<!-- PROFILE MODAL -->
<div class="mover" id="shareM">
  <div class="modal" style="max-width:340px">
    <div class="mhd"><span class="mttl">Share &amp; Earn</span><button class="mx" onclick="closeShare()">&#10005;</button></div>
    <div class="mbdy">
      <div style="font-size:11px;color:var(--dm);margin-bottom:14px">Invite friends to the dashboard — share your link below.</div>
      <div class="shareGrid">
        <div class="shareBtn" onclick="shareTo('whatsapp')"><span class="shIco" style="color:#25D366">&#9742;</span>WhatsApp</div>
        <div class="shareBtn" onclick="shareTo('twitter')"><span class="shIco" style="color:#1DA1F2">&#120143;</span>X / Twitter</div>
        <div class="shareBtn" onclick="shareTo('facebook')"><span class="shIco" style="color:#1877F2">f</span>Facebook</div>
        <div class="shareBtn" onclick="shareTo('linkedin')"><span class="shIco" style="color:#0A66C2">in</span>LinkedIn</div>
        <div class="shareBtn" onclick="shareTo('telegram')"><span class="shIco" style="color:#26A5E4">&#9992;</span>Telegram</div>
        <div class="shareBtn" onclick="shareTo('email')"><span class="shIco" style="color:var(--dm)">&#9993;</span>Email</div>
      </div>
      <div style="height:14px"></div>
      <label class="mlbl">Your Referral Link <span style="color:var(--dm)" id="refCountSh"></span></label>
      <div style="display:flex;gap:6px">
        <input class="minp" id="refLinkShare" type="text" style="height:34px;font-size:10px" readonly/>
        <button class="bok" style="padding:0 14px;flex-shrink:0" onclick="copyReferralLink()">Copy</button>
      </div>
    </div>
  </div>
</div>

<div class="mover" id="orderM">
  <div class="modal">
    <div class="mhd"><span class="mttl">Place Order</span><button class="mx" onclick="closeOrderModal()">&#10005;</button></div>
    <div class="mbdy">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px">
        <span style="font-size:18px;font-weight:700" id="obSym"></span>
        <span class="bdg ba" id="obCat"></span>
      </div>
      <div id="obAcctRow" style="display:none;margin-bottom:12px"><label class="mlbl">Account</label><div style="display:flex;gap:6px"><button class="bcx" id="obAcctMock" onclick="setOrderAcct('mock')">Mock</button><button class="bcx" id="obAcctLive" onclick="setOrderAcct('live')">Live (real money)</button></div></div>
      <label class="mlbl">Order Type</label>
      <select class="minp" id="obType" style="height:34px;margin-bottom:11px">
        <option value="BUY">Buy</option>
        <option value="SELL">Sell</option>
      </select>
      <label class="mlbl">Quantity</label>
      <input class="minp" id="obQty" type="number" style="height:34px;margin-bottom:11px"/>
      <div id="obKindRow" style="display:none"><label class="mlbl">Order kind (live)</label><select class="minp" id="obKind" style="height:34px;margin-bottom:11px" onchange="updOrderEstimate()"><option value="LIMIT">Limit</option><option value="MARKET">Market</option></select></div>
      <label class="mlbl">Price <span id="obLtpBadge" style="color:var(--ac);font-weight:600;text-transform:none;letter-spacing:0"></span></label>
      <input class="minp" id="obPrice" type="number" style="height:34px;margin-bottom:11px"/>
      <div id="obMockOnly">
      <label class="mlbl">Target (optional)</label>
      <input class="minp" id="obTarget" type="number" style="height:34px;margin-bottom:11px"/>
      <label class="mlbl">Stop Loss (optional)</label>
      <input class="minp" id="obSL" type="number" style="height:34px;margin-bottom:11px"/>
      <label class="mlbl">Notes (optional)</label>
      <input class="minp" id="obNotes" type="text" style="height:34px;margin-bottom:10px"/>
      </div>
      <div id="obEstimate" style="font-size:11px;color:var(--dm);margin-bottom:12px;min-height:16px"></div>
      <button class="bok" id="obSubmit" style="width:100%" onclick="submitOrder()">Place Order</button>
      <div id="obNote" style="font-size:10px;color:var(--dm);margin-top:10px;text-align:center">Orders are tracked in your Order Book. This does not place a live trade with any broker.</div>
    </div>
  </div>
</div>

<div class="mover" id="guideM">
  <div class="modal guideModal">
    <div class="mhd"><span class="mttl">Broker (Live) — setup guide</span><button class="mx" onclick="closeBrokerGuide()">&#10005;</button></div>
    <div class="mbdy guideBdy">
<div class="gd">
<p class="gdLead">This guide is only for <b>Live</b> trading. Mock trading needs none of it. About 45 minutes the first time.</p>

<div class="gdBox gdOk"><b>Does every user have to do all this?</b> Yes &mdash; for Live trading each person sets up <b>their own</b>: Kite Connect app, static IP (registered in their own Kite developer account) and relay. Zerodha accepts API orders only from the IP registered in the account's developer profile, and that IP may be used only by the account holder and their immediate family. <i>Family:</i> one developer account, one static IP and one relay can serve the whole family (a separate app per person) &mdash; each person still types their own API key, secret and user ID into their own Broker (Live) settings.<br><b>You do NOT repeat</b> the dashboard itself, any Script Properties, or the keep-alive &mdash; those belong to the dashboard owner and already cover every user's relay.</div>

<div class="gdBox"><b>Using the owner's relay instead of building your own.</b> If the dashboard owner has given you their relay address and relay secret you can skip Steps 2, 4, 5 and 6: do <b>Step 1</b> (your own Kite app) and <b>Step 7</b> with the relay URL and secret they gave you, then <b>Step 8</b>. Your orders still use <i>your</i> Zerodha credentials; the relay only passes them through, so only do this with someone you trust. There are two ways to handle the static IP:<br>&bull; <b>Immediate family</b> (Zerodha allows sharing an IP only with the holder and immediate family): do <b>Step 3</b> but whitelist <i>the owner's</i> static IP (they tell you) in <i>your</i> developer account, and leave "Your own static-IP proxy URL" empty.<br>&bull; <b>Anyone else</b>: you need your own static IP. Do <b>Steps 2 and 3</b> (your own TrueIP address, whitelisted in your developer account) and paste your proxy URL into <b>Broker (Live) &rarr; Your own static-IP proxy URL</b>. The owner's relay then sends your orders from your IP. Use <b>Test &mdash; show my IP</b> to confirm.</div>

<svg viewBox="0 0 700 170" class="gdSvg" role="img" aria-label="How the pieces connect">
<rect x="8" y="55" width="130" height="64" rx="8" fill="none" stroke="currentColor" opacity=".6"/><text x="73" y="82" text-anchor="middle" font-size="12" fill="currentColor">This dashboard</text><text x="73" y="100" text-anchor="middle" font-size="9" fill="currentColor" opacity=".7">(Apps Script)</text>
<rect x="188" y="55" width="150" height="64" rx="8" fill="none" stroke="currentColor" opacity=".6"/><text x="263" y="78" text-anchor="middle" font-size="12" fill="currentColor">Your relay</text><text x="263" y="94" text-anchor="middle" font-size="9" fill="currentColor" opacity=".8">Render (free)</text><text x="263" y="108" text-anchor="middle" font-size="9" fill="currentColor" opacity=".8">files from GitHub</text>
<rect x="388" y="55" width="150" height="64" rx="8" fill="rgba(0,230,138,.10)" stroke="#00e68a"/><text x="463" y="78" text-anchor="middle" font-size="12" fill="currentColor">TrueIP proxy</text><text x="463" y="94" text-anchor="middle" font-size="9" fill="#00e68a">your STATIC IP</text><text x="463" y="108" text-anchor="middle" font-size="9" fill="#00e68a">whitelisted on Zerodha &#10003;</text>
<rect x="588" y="55" width="104" height="64" rx="8" fill="none" stroke="currentColor" opacity=".6"/><text x="640" y="82" text-anchor="middle" font-size="12" fill="currentColor">Zerodha</text><text x="640" y="100" text-anchor="middle" font-size="9" fill="currentColor" opacity=".7">api.kite.trade</text>
<path d="M138 87h50M338 87h50M538 87h50" stroke="currentColor" stroke-width="2" marker-end="url(#gar)"/>
<text x="163" y="80" text-anchor="middle" font-size="8" fill="currentColor" opacity=".7">orders</text><text x="363" y="80" text-anchor="middle" font-size="8" fill="currentColor" opacity=".7">orders</text><text x="563" y="80" text-anchor="middle" font-size="8" fill="currentColor" opacity=".7">orders</text>
<path d="M73 55 C73 12 640 12 640 55" fill="none" stroke="currentColor" stroke-dasharray="4 4" opacity=".6"/><text x="356" y="24" text-anchor="middle" font-size="9" fill="currentColor" opacity=".8">positions, holdings, margins &mdash; direct, no static IP needed</text>
<defs><marker id="gar" markerWidth="8" markerHeight="8" refX="6" refY="3" orient="auto"><path d="M0 0L6 3L0 6z" fill="currentColor"/></marker></defs>
<text x="350" y="160" text-anchor="middle" font-size="9" fill="currentColor" opacity=".6">Illustration</text>
</svg>

<h4>What you need (accounts)</h4>
<table class="gdTbl"><thead><tr><th>Account</th><th>Cost</th><th>Used for</th></tr></thead><tbody>
<tr><td>Zerodha trading account</td><td>you have it</td><td>the account that trades</td></tr>
<tr><td>Kite Connect developer account &mdash; <b>developers.kite.trade</b></td><td>Free &ldquo;Personal&rdquo; plan is enough</td><td>API key + secret, IP whitelist</td></tr>
<tr><td><b>TrueIP</b> (static IP)</td><td>Paid &mdash; check TrueIP's current price</td><td>the fixed IP Zerodha will see</td></tr>
<tr><td><b>GitHub</b></td><td>Free</td><td>stores the 2 relay files</td></tr>
<tr><td><b>Render.com</b></td><td>Free instance</td><td>runs the relay 24/7</td></tr>
</tbody></table>
<div class="gdBox"><b>Your redirect URL</b> (needed in Step 1):<br><code id="gdRedirect">loading&hellip;</code> <button class="bcx" onclick="copyGuideRedirect(this)" style="padding:2px 8px">Copy</button></div>

<h4>Step 1 &mdash; Create your Kite Connect app</h4>
<ol>
<li>Open <b>developers.kite.trade</b>, sign in with your Zerodha ID, then <b>My apps &rarr; Create new app</b>.</li>
<li>Choose <b>Personal</b> (free; orders, positions, holdings, funds &mdash; no market data, which this dashboard doesn't need).</li>
<li>Fill the form: <b>App name</b> (anything), <b>Zerodha Client ID</b> (your user ID, like <code>AB1234</code>), <b>Redirect URL</b> = the URL in the box above, <b>Postback URL</b> = leave empty.</li>
<li>Create the app, open it, and copy the <b>API key</b> and <b>API secret</b>. Keep the secret private.</li>
</ol>
<svg viewBox="0 0 640 225" class="gdSvg" role="img" aria-label="Create new app form">
<rect x="20" y="8" width="600" height="209" rx="10" fill="none" stroke="currentColor" opacity=".5"/>
<text x="40" y="34" font-size="13" fill="currentColor" font-weight="700">Create new app</text>
<text x="40" y="58" font-size="10" fill="currentColor" opacity=".7">Type</text><rect x="40" y="64" width="250" height="24" rx="4" fill="none" stroke="currentColor" opacity=".5"/><text x="50" y="80" font-size="11" fill="currentColor">Personal (free)</text>
<text x="330" y="58" font-size="10" fill="currentColor" opacity=".7">App name</text><rect x="330" y="64" width="270" height="24" rx="4" fill="none" stroke="currentColor" opacity=".5"/><text x="340" y="80" font-size="11" fill="currentColor">My Signal Dashboard</text>
<text x="40" y="110" font-size="10" fill="currentColor" opacity=".7">Zerodha Client ID</text><rect x="40" y="116" width="250" height="24" rx="4" fill="none" stroke="currentColor" opacity=".5"/><text x="50" y="132" font-size="11" fill="currentColor">AB1234</text>
<text x="330" y="110" font-size="10" fill="currentColor" opacity=".7">Postback URL</text><rect x="330" y="116" width="270" height="24" rx="4" fill="none" stroke="currentColor" opacity=".5" stroke-dasharray="3 3"/><text x="340" y="132" font-size="11" fill="currentColor" opacity=".6">(leave empty)</text>
<text x="40" y="162" font-size="10" fill="currentColor" opacity=".7">Redirect URL</text><rect x="40" y="168" width="560" height="24" rx="4" fill="rgba(0,230,138,.10)" stroke="#00e68a"/><text x="50" y="184" font-size="11" fill="currentColor">https://script.google.com/macros/s/&hellip;/exec</text>
<text x="40" y="208" font-size="9" fill="currentColor" opacity=".6">Illustration &mdash; Zerodha's real page may look slightly different</text>
</svg>

<h4>Step 2 &mdash; Get a static IP from TrueIP</h4>
<ol>
<li>Create an account at TrueIP (the provider the dashboard owner uses) and buy a <b>dedicated static IP</b> plan.</li>
<li>In your TrueIP dashboard (or the CSV it lets you export) note two things: <b>(a)</b> your <b>static IP address</b> &mdash; this is what Zerodha will see, and <b>(b)</b> the proxy details: <b>username, password, host</b> (<code>gw.trueip.in</code>) and <b>port</b>.</li>
<li>Join them into one line &mdash; this is your <b>proxy URL</b> (you'll paste it into Render in Step 5):<br><code>http://USERNAME:PASSWORD@gw.trueip.in:PORT</code><br>If the username or password contains <code>@ : / # ? %</code>, replace each with its code: <code>@</code>&rarr;<code>%40</code> &nbsp; <code>:</code>&rarr;<code>%3A</code> &nbsp; <code>/</code>&rarr;<code>%2F</code> &nbsp; <code>#</code>&rarr;<code>%23</code> &nbsp; <code>?</code>&rarr;<code>%3F</code> &nbsp; <code>%</code>&rarr;<code>%25</code>.</li>
</ol>
<p class="gdNote">I couldn't view TrueIP's own screens, so menu names may differ &mdash; ask the dashboard owner or TrueIP support if something isn't where this says. Any provider that sells a dedicated static IP with an HTTP proxy works the same way.</p>

<h4>Step 3 &mdash; Whitelist the IP on Zerodha</h4>
<ol>
<li>On <b>developers.kite.trade</b> open <b>Profile</b> (top right) &rarr; <b>IP Whitelist</b>.</li>
<li>Enter your static IP as the <b>primary</b> IP (a secondary is optional), confirm the declaration that it is used only by you / your immediate family, and press <b>Update</b>.</li>
<li><b>Check it character by character first</b> &mdash; Zerodha allows only <b>one change per calendar week</b>.</li>
</ol>
<svg viewBox="0 0 640 110" class="gdSvg" role="img" aria-label="IP whitelist">
<rect x="20" y="8" width="600" height="94" rx="10" fill="none" stroke="currentColor" opacity=".5"/>
<text x="40" y="32" font-size="12" fill="currentColor" font-weight="700">Profile &rarr; IP Whitelist</text>
<text x="40" y="54" font-size="10" fill="currentColor" opacity=".7">Primary static IP</text><rect x="40" y="60" width="260" height="26" rx="4" fill="rgba(0,230,138,.10)" stroke="#00e68a"/><text x="50" y="78" font-size="12" fill="currentColor">203.0.113.10</text>
<rect x="320" y="60" width="80" height="26" rx="4" fill="#00e68a"/><text x="360" y="78" text-anchor="middle" font-size="11" fill="#000" font-weight="700">Update</text>
<text x="420" y="78" font-size="10" fill="currentColor" opacity=".7">&larr; your TrueIP address</text>
<text x="620" y="98" text-anchor="end" font-size="9" fill="currentColor" opacity=".6">Illustration (example IP)</text>
</svg>

<h4>Step 4 &mdash; Put the relay files on GitHub</h4>
<ol>
<li>Create a free account at <b>github.com</b> (skip if you have one).</li>
<li>Click <b>+ &rarr; New repository</b>, name it <code>kite-relay</code>, leave the defaults, press <b>Create repository</b>.</li>
<li>Save the files below (use <b>Download</b> or <b>Copy</b>), then in the new repository press <b>Add file &rarr; Upload files</b>, drop <code>index.js</code> and <code>package.json</code> (and optionally <code>.env.example</code>), and press <b>Commit changes</b>.</li>
</ol>
<p class="gdFile"><b>File 1 of 3 &mdash; <code>index.js</code></b> (required) &mdash; the relay itself. It sends orders to Zerodha through the TrueIP proxy (yours by default, or a person's own proxy if they supply one), forwards <code>market_protection</code> (required on market orders), and can change / cancel pending orders.</p>
<div class="gdCode"><div class="gdBtns"><button class="bcx" onclick="copyGuideBlock(this,'gdFileIndex')">Copy</button><button class="bcx" onclick="downloadGuideBlock(this,'gdFileIndex','index.js')">Download</button></div><pre id="gdFileIndex">// Kite order relay (v5) - sends orders to Zerodha from a STATIC IP, through an HTTP proxy (TrueIP or any provider with the same
// host:port:user:pass model). Zerodha accepts API orders only from the IP registered in the account's developer profile.
//
// v5: ONE relay can serve many people. A request may carry its own `proxy_url` (that person's own static-IP proxy); without one the
// relay uses STATIC_IP_PROXY_URL (the relay owner's own IP). Adds /modify-order, /cancel-order and /check-proxy.
//
// Security: every request must include RELAY_SECRET. The relay never stores or logs api keys, access tokens or proxy URLs.

const express = require('express');
const https = require('https');
const crypto = require('crypto');
const { HttpsProxyAgent } = require('https-proxy-agent');

const app = express();
app.use(express.json({ limit: '50kb' }));

const PORT = process.env.PORT || 8080;
const RELAY_SECRET = process.env.RELAY_SECRET;
const PROXY_URL = process.env.STATIC_IP_PROXY_URL; // optional in v5: http://USER:PASS@gw.trueip.in:PORT (the owner's own IP)

if (!RELAY_SECRET) {
  console.error('RELAY_SECRET is not set. Refusing to start - this relay would accept unauthenticated order requests otherwise.');
  process.exit(1);
}
if (!PROXY_URL) {
  console.warn('STATIC_IP_PROXY_URL is not set. Requests that do not bring their own proxy_url will be refused (they would leave from this server\'s own IP, which Kite rejects).');
}

const defaultAgent = PROXY_URL ? new HttpsProxyAgent(PROXY_URL) : null;
const agentCache = new Map(); // proxy_url -&gt; agent (small LRU)

function secretOk(given) {
  const a = Buffer.from(String(given || '')), b = Buffer.from(String(RELAY_SECRET));
  return a.length === b.length &amp;&amp; crypto.timingSafeEqual(a, b);
}

// A caller-supplied proxy must be a plain http(s) proxy on a PUBLIC host - never the relay's own network.
function isPrivateHost(h) {
  h = String(h || '').toLowerCase().replace(/^\[|\]$/g, '');
  if (!h || h === 'localhost' || h.endsWith('.local') || h.endsWith('.internal') || h.endsWith('.localhost')) return true;
  if (/^(::1|::|fc|fd|fe[89ab])/i.test(h) &amp;&amp; h.includes(':')) return true;
  const m = /^(\d+)\.(\d+)\.(\d+)\.(\d+)$/.exec(h);
  if (m) {
    const [a, b] = [Number(m[1]), Number(m[2])];
    if (a === 0 || a === 10 || a === 127 || (a === 169 &amp;&amp; b === 254) || (a === 172 &amp;&amp; b &gt;= 16 &amp;&amp; b &lt;= 31) || (a === 192 &amp;&amp; b === 168) || a &gt;= 224) return true;
  }
  return false;
}
function agentFor(proxyUrl) {
  if (!proxyUrl) {
    if (!defaultAgent) throw new Error('This relay has no default proxy and the request did not include proxy_url.');
    return defaultAgent;
  }
  const s = String(proxyUrl);
  if (s.length &gt; 500) throw new Error('proxy_url is too long.');
  let u;
  try { u = new URL(s); } catch (e) { throw new Error('proxy_url is not a valid URL.'); }
  if (u.protocol !== 'http:' &amp;&amp; u.protocol !== 'https:') throw new Error('proxy_url must start with http:// or https://');
  if (isPrivateHost(u.hostname)) throw new Error('proxy_url must point to a public proxy host.');
  const port = Number(u.port || (u.protocol === 'https:' ? 443 : 80));
  if (!(port &gt;= 1 &amp;&amp; port &lt;= 65535)) throw new Error('proxy_url has an invalid port.');
  let ag = agentCache.get(s);
  if (!ag) {
    ag = new HttpsProxyAgent(s);
    agentCache.set(s, ag);
    if (agentCache.size &gt; 50) agentCache.delete(agentCache.keys().next().value);
  }
  return ag;
}

function log(body, outcome) {
  console.log(new Date().toISOString(), outcome, (body &amp;&amp; body.tradingsymbol) || (body &amp;&amp; body.order_id) || '', (body &amp;&amp; body.transaction_type) || '', (body &amp;&amp; body.quantity) || '', (body &amp;&amp; body.order_type) || '');
}

// One HTTPS request to api.kite.trade through the chosen proxy.
function kiteRequest(method, path, params, headers, agent) {
  return new Promise((resolve, reject) =&gt; {
    const payload = params ? params.toString() : '';
    const h = Object.assign({}, headers);
    if (payload) { h['Content-Type'] = 'application/x-www-form-urlencoded'; h['Content-Length'] = Buffer.byteLength(payload); }
    const req = https.request({ hostname: 'api.kite.trade', path, method, agent, headers: h }, (res) =&gt; {
      let body = '';
      res.on('data', (c) =&gt; { body += c; });
      res.on('end', () =&gt; {
        try { resolve({ statusCode: res.statusCode, data: JSON.parse(body) }); }
        catch (e) { reject(new Error('Kite returned a non-JSON response: ' + body.slice(0, 200))); }
      });
    });
    req.on('error', reject);
    if (payload) req.write(payload);
    req.end();
  });
}

function authHeaders(body) {
  return { 'X-Kite-Version': '3', 'Authorization': `token ${body.api_key}:${body.access_token}` };
}
function pick(body, keys) {
  const f = new URLSearchParams();
  keys.forEach((k) =&gt; { if (body[k] !== undefined &amp;&amp; body[k] !== null &amp;&amp; body[k] !== '') f.append(k, body[k]); });
  return f;
}
function guard(req, res, needOrderId) {
  const body = req.body || {};
  if (!secretOk(body.relay_secret)) { log(body, 'REJECTED (bad secret)'); res.status(403).json({ success: false, error: 'Invalid relay secret.' }); return null; }
  if (!body.api_key || !body.access_token) { res.status(400).json({ success: false, error: 'Missing api_key or access_token.' }); return null; }
  if (needOrderId &amp;&amp; !/^[0-9A-Za-z_-]{6,40}$/.test(String(body.order_id || ''))) { res.status(400).json({ success: false, error: 'Missing or invalid order_id.' }); return null; }
  let agent;
  try { agent = agentFor(body.proxy_url); } catch (e) { res.status(400).json({ success: false, error: e.message }); return null; }
  return { body, agent };
}
function sendKite(res, body, kiteData, okLabel) {
  if (kiteData.status !== 'success') {
    log(body, 'REJECTED by Kite: ' + kiteData.error_type);
    return res.json({ success: false, error: kiteData.message || 'Kite rejected the request.', error_type: kiteData.error_type });
  }
  log(body, okLabel + ' ' + (kiteData.data &amp;&amp; kiteData.data.order_id));
  return res.json({ success: true, order_id: kiteData.data &amp;&amp; kiteData.data.order_id });
}
const VARIETIES = ['regular', 'amo', 'co', 'iceberg', 'auction'];
const variety = (b) =&gt; (VARIETIES.includes(String(b.variety || 'regular').toLowerCase()) ? String(b.variety || 'regular').toLowerCase() : 'regular');

// ---- place ----
app.post('/place-order', async (req, res) =&gt; {
  const g = guard(req, res, false); if (!g) return;
  const body = g.body;
  if (!body.tradingsymbol || !body.transaction_type || !body.quantity || !body.order_type) {
    return res.status(400).json({ success: false, error: 'Missing tradingsymbol, transaction_type, quantity, or order_type.' });
  }
  // Everything Kite's order endpoint accepts, forwarded as-is (market_protection is required by Kite on API MARKET orders).
  const params = pick(body, ['tradingsymbol', 'exchange', 'transaction_type', 'order_type', 'quantity', 'product', 'validity', 'price', 'trigger_price', 'disclosed_quantity', 'squareoff', 'stoploss', 'trailing_stoploss', 'tag', 'market_protection']);
  try {
    const { data } = await kiteRequest('POST', '/orders/regular', params, authHeaders(body), g.agent);
    return sendKite(res, body, data, 'PLACED');
  } catch (err) {
    console.error(new Date().toISOString(), 'RELAY ERROR', err.message);
    return res.status(502).json({ success: false, error: 'Relay could not reach Kite through the proxy: ' + err.message });
  }
});

// ---- modify a pending order ----
app.post('/modify-order', async (req, res) =&gt; {
  const g = guard(req, res, true); if (!g) return;
  const body = g.body;
  const params = pick(body, ['order_type', 'quantity', 'price', 'trigger_price', 'validity', 'disclosed_quantity']);
  try {
    const { data } = await kiteRequest('PUT', `/orders/${variety(body)}/${body.order_id}`, params, authHeaders(body), g.agent);
    return sendKite(res, body, data, 'MODIFIED');
  } catch (err) {
    console.error(new Date().toISOString(), 'RELAY ERROR', err.message);
    return res.status(502).json({ success: false, error: 'Relay could not reach Kite through the proxy: ' + err.message });
  }
});

// ---- cancel a pending order ----
app.post('/cancel-order', async (req, res) =&gt; {
  const g = guard(req, res, true); if (!g) return;
  const body = g.body;
  try {
    const { data } = await kiteRequest('DELETE', `/orders/${variety(body)}/${body.order_id}`, null, authHeaders(body), g.agent);
    return sendKite(res, body, data, 'CANCELLED');
  } catch (err) {
    console.error(new Date().toISOString(), 'RELAY ERROR', err.message);
    return res.status(502).json({ success: false, error: 'Relay could not reach Kite through the proxy: ' + err.message });
  }
});

function egressIp(agent) {
  return new Promise((resolve, reject) =&gt; {
    const r = https.request({ hostname: 'ifconfig.me', path: '/ip', method: 'GET', agent }, (resp) =&gt; {
      let b = ''; resp.on('data', (c) =&gt; (b += c)); resp.on('end', () =&gt; resolve(b.trim()));
    });
    r.on('error', reject); r.end();
  });
}

// ---- which IP does a proxy leave from? (so a user can check it matches the IP whitelisted on Zerodha) ----
app.post('/check-proxy', async (req, res) =&gt; {
  const body = req.body || {};
  if (!secretOk(body.relay_secret)) return res.status(403).json({ ok: false, error: 'Invalid relay secret.' });
  try {
    const ip = await egressIp(agentFor(body.proxy_url));
    return res.json({ ok: true, egressIp: ip, usingOwnProxy: !!body.proxy_url });
  } catch (e) {
    return res.json({ ok: false, error: 'Proxy check failed: ' + e.message });
  }
});

// Confirms the relay is alive AND that its DEFAULT proxy works (the IP shown must be the one you whitelisted with Zerodha).
// Also a keep-alive target for hosts that put idle services to sleep (Render's free tier).
app.get('/health', async (req, res) =&gt; {
  if (!defaultAgent) return res.json({ ok: true, time: new Date().toISOString(), defaultProxy: false });
  try {
    res.json({ ok: true, time: new Date().toISOString(), egressIpThroughProxy: await egressIp(defaultAgent) });
  } catch (e) {
    res.status(500).json({ ok: false, error: 'Relay is up but the proxy leg failed: ' + e.message });
  }
});

app.listen(PORT, () =&gt; console.log(`Kite order relay (v5) listening on port ${PORT}${PROXY_URL ? ', default proxy configured' : ', no default proxy (callers bring their own)'}`));</pre></div>
<p class="gdFile"><b>File 2 of 3 &mdash; <code>package.json</code></b> (required) &mdash; tells Render which libraries to install.</p>
<div class="gdCode"><div class="gdBtns"><button class="bcx" onclick="copyGuideBlock(this,'gdFilePkg')">Copy</button><button class="bcx" onclick="downloadGuideBlock(this,'gdFilePkg','package.json')">Download</button></div><pre id="gdFilePkg">{
  "name": "kite-order-relay",
  "version": "2.0.0",
  "description": "Relay that lets Apps Script place real Kite Connect orders, routed through a static-IP proxy (TrueIP or similar) since that's what Kite requires for order placement specifically.",
  "main": "index.js",
  "scripts": {
    "start": "node index.js"
  },
  "engines": {
    "node": "&gt;=18"
  },
  "dependencies": {
    "express": "^4.19.2",
    "https-proxy-agent": "^9.1.0"
  }
}</pre></div>
<p class="gdFile"><b>File 3 of 3 &mdash; <code>.env.example</code></b> (optional) &mdash; only a reminder of the two settings Render needs. It holds <b>no real values</b>; never put your real secret or proxy URL in a file on GitHub.</p>
<div class="gdCode"><div class="gdBtns"><button class="bcx" onclick="copyGuideBlock(this,'gdFileEnv')">Copy</button><button class="bcx" onclick="downloadGuideBlock(this,'gdFileEnv','.env.example')">Download</button></div><pre id="gdFileEnv"># Copy this file to .env and fill in your own values — never commit the real .env to git.

# A random string only this relay and your dashboard (Profile &gt; Broker (Live) &gt; Relay secret) know.
# Generate one with: node -e "console.log(require('crypto').randomBytes(24).toString('hex'))"
RELAY_SECRET=

# Your static-IP proxy's full URL, credentials included — e.g. from TrueIP's dashboard after
# signup: http://USER:PASS@gw.trueip.in:PORT (host/port/user/pass are per-account, shown on your
# TrueIP dashboard or exported CSV). This is the address that actually gets whitelisted with
# Zerodha, NOT this server's own IP.
STATIC_IP_PROXY_URL=

# Optional — defaults to 8080 if not set.
PORT=8080</pre></div>
<div class="gdBox"><b>Already running a relay from before?</b> Replace your whole <code>index.js</code> with File 1 above: on GitHub open <code>index.js</code> &rarr; pencil icon &rarr; select all &rarr; paste &rarr; <b>Commit changes</b>. Render redeploys by itself. The new version also lets people bring their own proxy, changes/cancels pending orders, and checks which IP a proxy uses.</div>

<h4>Step 5 &mdash; Run it on Render</h4>
<ol>
<li>Go to <b>render.com</b> &rarr; <b>Get Started</b> &rarr; <b>Sign in with GitHub</b> and approve.</li>
<li><b>New + &rarr; Web Service</b> &rarr; pick your <code>kite-relay</code> repository &rarr; <b>Connect</b>.</li>
<li>Settings: <b>Name</b> anything (e.g. <code>kite-relay</code>) &middot; <b>Language</b> Node &middot; <b>Build Command</b> <code>npm install</code> &middot; <b>Start Command</b> <code>node index.js</code> &middot; <b>Instance Type</b> Free.</li>
<li>Under <b>Environment Variables</b> add two (names must match exactly):<br><code>RELAY_SECRET</code> = a long random text you invent (20&ndash;30 characters) &mdash; <b>save it, you'll need it again in Step 7</b><br><code>STATIC_IP_PROXY_URL</code> = the proxy URL from Step 2 (optional if everyone using this relay brings their own proxy URL)</li>
<li>Press <b>Create Web Service</b>. After 1&ndash;3 minutes it says <b>Live</b> and shows an address like <code>https://kite-relay-xxxx.onrender.com</code>. Copy it.</li>
</ol>
<svg viewBox="0 0 640 130" class="gdSvg" role="img" aria-label="Render environment variables">
<rect x="20" y="8" width="600" height="114" rx="10" fill="none" stroke="currentColor" opacity=".5"/>
<text x="40" y="32" font-size="12" fill="currentColor" font-weight="700">Environment Variables</text>
<text x="40" y="52" font-size="9" fill="currentColor" opacity=".7">NAME</text><text x="260" y="52" font-size="9" fill="currentColor" opacity=".7">VALUE</text>
<rect x="40" y="58" width="200" height="22" rx="4" fill="none" stroke="currentColor" opacity=".5"/><text x="50" y="73" font-size="11" fill="currentColor">RELAY_SECRET</text>
<rect x="260" y="58" width="340" height="22" rx="4" fill="rgba(0,230,138,.10)" stroke="#00e68a"/><text x="270" y="73" font-size="11" fill="currentColor">a-long-random-text-you-invent</text>
<rect x="40" y="88" width="200" height="22" rx="4" fill="none" stroke="currentColor" opacity=".5"/><text x="50" y="103" font-size="11" fill="currentColor">STATIC_IP_PROXY_URL</text>
<rect x="260" y="88" width="340" height="22" rx="4" fill="rgba(0,230,138,.10)" stroke="#00e68a"/><text x="270" y="103" font-size="11" fill="currentColor">http://USER:PASS@gw.trueip.in:PORT</text>
<text x="620" y="120" text-anchor="end" font-size="9" fill="currentColor" opacity=".6">Illustration</text>
</svg>

<h4>Step 6 &mdash; Test the relay</h4>
<p>Open <code>https://kite-relay-xxxx.onrender.com/health</code> (your address + <code>/health</code>) in a browser. You should see <code>{"ok":true,"time":"&hellip;","egressIpThroughProxy":"&hellip;"}</code>. <b><code>egressIpThroughProxy</code> must be exactly the IP you whitelisted in Step 3.</b> If it isn't, or you see an error, stop and fix it before going on (see the table at the end).</p>

<h4>Step 7 &mdash; Fill Profile &rarr; Broker (Live)</h4>
<table class="gdTbl"><thead><tr><th>Field</th><th>What to enter</th><th>Where it comes from</th></tr></thead><tbody>
<tr><td>Kite Connect API key</td><td>the key of your app</td><td>Step 1</td></tr>
<tr><td>API secret</td><td>the secret of your app</td><td>Step 1</td></tr>
<tr><td>Zerodha user ID</td><td>e.g. <code>AB1234</code></td><td>your Zerodha profile</td></tr>
<tr><td>Order relay URL</td><td><code>https://kite-relay-xxxx.onrender.com/place-order</code></td><td>Step 5 address + <code>/place-order</code></td></tr>
<tr><td>Relay secret</td><td>exactly the text you used for <code>RELAY_SECRET</code></td><td>Step 5 (no extra spaces)</td></tr>
<tr><td>Your own static-IP proxy URL <i>(optional)</i></td><td><code>http://USER:PASS@gw.trueip.in:PORT</code></td><td>Step 2 &mdash; only if you use someone else's relay with your own IP</td></tr>
</tbody></table>
<p>Press <b>Save Broker Settings</b>. (If two-step verification is on, you'll be asked for a code.) Leave the password and TOTP boxes empty unless you want the optional automatic login below.</p>

<h4>Step 8 &mdash; Connect, then test small</h4>
<ol>
<li>Open <b>Portfolio</b>, switch to <b>Live</b>, press <b>Connect Zerodha</b>, sign in on Zerodha's page and approve. The status turns <b>Connected</b>.</li>
<li>Zerodha's connection lasts one trading day, so repeat each morning &mdash; or use the optional automatic login.</li>
<li>Before trusting it, press <b>+ Order</b> (or the order button on any stock in the Live Panel, choosing <b>Live</b>) and place a <b>1-share limit order</b> on a cheap, liquid stock. It appears in the <b>Live</b> side of your <b>Order Book</b>, where you can also change or cancel it while it is pending, and in your Zerodha order book.</li>
<li>The free Render instance falls asleep when idle. You don't need to do anything: the dashboard pings every user's relay every 10 minutes (the owner's keep-alive).</li>
</ol>

<h4>Optional &mdash; Automatic daily login (needs your TOTP secret)</h4>
<div class="gdBox gdOk"><b>TOTP is NOT required.</b> Skip this section and just press <i>Connect Zerodha</i> each morning. Orders, positions and everything else work without it.</div>
<ul>
<li>It needs your Zerodha <b>password</b> and <b>TOTP secret key</b>. The key exists only if your Zerodha login uses <b>External authenticator (TOTP)</b>, not the &ldquo;Kite app code&rdquo; method. Switching changes how you log in to Zerodha yourself, too.</li>
<li>To get it: Kite on the web &rarr; click your name/ID (top right) &rarr; profile / settings &rarr; <b>Security</b> (2FA) &rarr; enable <b>External TOTP</b>. On the QR screen choose <b>&ldquo;Can't scan? Copy key&rdquo;</b> and paste that key into the <i>TOTP secret</i> box (also scan the QR in your authenticator app so you can still log in yourself). Zerodha shows the key once &mdash; if you missed it, turn TOTP off and on again.</li>
<li>Zerodha does not provide or endorse automated login and it can stop working at any time; your password and key are stored (encrypted) in this app's Apps Script project. Use it only if you accept that.</li>
</ul>
<svg viewBox="0 0 640 120" class="gdSvg" role="img" aria-label="Where to find the TOTP key">
<rect x="20" y="8" width="600" height="100" rx="10" fill="none" stroke="currentColor" opacity=".5"/>
<text x="40" y="32" font-size="12" fill="currentColor" font-weight="700">Enable external TOTP</text>
<rect x="40" y="44" width="70" height="56" fill="none" stroke="currentColor" opacity=".5"/><text x="75" y="76" text-anchor="middle" font-size="10" fill="currentColor" opacity=".7">QR code</text>
<text x="130" y="60" font-size="10" fill="currentColor" opacity=".8">Scan with your authenticator app</text>
<rect x="130" y="70" width="250" height="24" rx="4" fill="rgba(0,230,138,.10)" stroke="#00e68a"/><text x="140" y="86" font-size="11" fill="currentColor">Can't scan? Copy key &rarr;  ABCD EFGH IJKL &hellip;</text>
<text x="400" y="86" font-size="10" fill="#00e68a">&larr; this is the TOTP secret</text>
<text x="620" y="116" text-anchor="end" font-size="9" fill="currentColor" opacity=".6">Illustration</text>
</svg>

<h4>If something goes wrong</h4>
<table class="gdTbl"><thead><tr><th>You see</th><th>What it means / what to do</th></tr></thead><tbody>
<tr><td>&ldquo;Add your static-IP order relay URL&hellip;&rdquo;</td><td>The Order relay URL box (Step 7) is empty.</td></tr>
<tr><td>&ldquo;Could not reach the order relay&rdquo;</td><td>The Render service is asleep, stopped, or the URL is wrong. Open <code>/health</code> in a browser; the first hit after a long idle can take about a minute.</td></tr>
<tr><td><code>/health</code> shows an error</td><td>Usually a typo in <code>STATIC_IP_PROXY_URL</code> or a wrong TrueIP username/password/port. Fix it in Render &rarr; Environment (it restarts by itself). Remember to code special characters (Step 2).</td></tr>
<tr><td><code>egressIpThroughProxy</code> is not your whitelisted IP</td><td>You whitelisted a different IP than TrueIP's static IP, or the proxy URL belongs to another plan. Correct the IP in Zerodha (one change per week &mdash; double-check) or the proxy URL in Render.</td></tr>
<tr><td>Zerodha says the IP is not allowed / not whitelisted</td><td>Same as above, or the whitelist change hasn't taken effect yet &mdash; connect again (Step 8) and retry.</td></tr>
<tr><td>&ldquo;Invalid relay secret.&rdquo;</td><td>The Relay secret in Settings differs from <code>RELAY_SECRET</code> in Render. Re-type both, no spaces.</td></tr>
<tr><td>&ldquo;Market orders without market protection&rdquo;</td><td>Your relay's <code>index.js</code> is the old one. Use File 1 above (see the grey box in Step 4). The dashboard also retries once as a protected limit order.</td></tr>
<tr><td>TokenException / &ldquo;Not connected&rdquo;</td><td>Zerodha's login has expired (one day). Press <i>Connect Zerodha</i> again.</td></tr>
<tr><td>MarginException</td><td>Not enough funds or margin in your Zerodha account for that order.</td></tr>
</tbody></table>
</div>
    </div>
  </div>
</div>
<div class="mover" id="profM">
  <div class="modal">
    <div class="mhd"><span class="mttl">Settings</span><button class="mx" onclick="closeProfile()">&#10005;</button></div>
    <div class="mbdy">
      <div class="authErr" id="profErr"></div>
      <div class="authMsg" id="profMsg"></div>

      <details class="setSec">
        <summary>Profile</summary>
        <div class="setBody">
          <label class="mlbl">Display Name</label>
          <input class="minp" id="profName" type="text" style="height:34px"/>
          <button class="bok" style="margin:8px 0 14px;width:100%" onclick="saveProfileName()">Save Name</button>
          <label class="mlbl">Email</label>
          <input class="minp" id="profEmail" type="text" style="height:34px" disabled/>
        </div>
      </details>

      <details class="setSec">
        <summary>Trading Preferences</summary>
        <div class="setBody">
          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">
            <span style="font-size:11px;color:var(--tx)">Execution Mode</span>
            <div class="mtyps" id="execModeBtns" style="margin-bottom:0">
              <button class="tybtn" data-mode="manual" onclick="setExecMode('manual')">Manual</button>
              <button class="tybtn" data-mode="auto" onclick="setExecMode('auto')">Auto</button>
            </div>
          </div>
          <div style="font-size:10px;color:var(--dm);margin-bottom:12px;line-height:1.5">In Auto mode, the AI agent may place, exit, or average mock trades on your behalf for the timeframes enabled below. This is a paper-trading playground — it never touches real money or a real broker.</div>
          <label class="mlbl">Margin (mock trading capital)</label>
          <div style="font-size:10px;color:var(--dm);margin-bottom:6px">Already invested across your whole portfolio: <b id="profInvested" style="color:var(--tx)">-</b></div>
          <div id="profMarginUsed" style="font-size:10px;color:var(--dm);margin-bottom:6px"></div>
          <div style="display:flex;gap:6px;margin-bottom:12px">
            <input class="minp" id="profMargin" type="number" placeholder="e.g. 100000" style="height:34px"/>
            <button class="bok" style="padding:0 14px;flex-shrink:0" onclick="saveMargin()">Save</button>
          </div>
          <label class="mlbl">Per-trade Cap (max MARGIN the AI uses on one trade, &#8377; — with margin privilege on, an intraday trade's value can be this &times; your leverage)</label>
          <div style="display:flex;gap:6px;margin-bottom:12px">
            <input class="minp" id="profCap" type="number" placeholder="e.g. 10000" style="height:34px"/>
            <button class="bok" style="padding:0 14px;flex-shrink:0" onclick="saveCap()">Save</button>
          </div>
          <label class="mlbl">Auto-trade by timeframe</label>
          <label class="authChk" style="margin-bottom:8px"><input type="checkbox" id="autoIntraday"/> Intraday</label>
          <label class="authChk" style="margin-bottom:8px"><input type="checkbox" id="autoShortterm"/> Short-term</label>
          <label class="authChk" style="margin-bottom:8px"><input type="checkbox" id="autoLongterm"/> Long-term</label>
          <button class="bok" style="width:100%" onclick="saveAutoToggles()">Save Auto-trade Preferences</button>

          <div style="border-top:1px solid var(--bd);margin:16px 0 12px"></div>
          <label class="mlbl">Contribution % by Timeframe</label>
          <div style="font-size:10px;color:var(--dm);margin-bottom:10px;line-height:1.5">When set above 0%, AI trades in that timeframe are sized as this % of your total margin (capped by whatever margin is still free) together with the per-trade cap above (whichever is smaller wins). Leave at 0 to keep using the per-trade cap for that timeframe.</div>
          <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px">
            <span style="font-size:10px;color:var(--dm);min-width:72px">Intraday</span>
            <input type="range" id="profContribIntraday" min="0" max="100" step="1" value="0" style="flex:1" oninput="document.getElementById('contribIntradayLbl').textContent=fmtContribPct_(this.value)"/>
            <span id="contribIntradayLbl" style="font-size:11px;color:var(--tx);min-width:64px;text-align:right">Cap</span>
          </div>
          <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px">
            <span style="font-size:10px;color:var(--dm);min-width:72px">Short-term</span>
            <input type="range" id="profContribShortterm" min="0" max="100" step="1" value="0" style="flex:1" oninput="document.getElementById('contribShorttermLbl').textContent=fmtContribPct_(this.value)"/>
            <span id="contribShorttermLbl" style="font-size:11px;color:var(--tx);min-width:64px;text-align:right">Cap</span>
          </div>
          <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px">
            <span style="font-size:10px;color:var(--dm);min-width:72px">Long-term</span>
            <input type="range" id="profContribLongterm" min="0" max="100" step="1" value="0" style="flex:1" oninput="document.getElementById('contribLongtermLbl').textContent=fmtContribPct_(this.value)"/>
            <span id="contribLongtermLbl" style="font-size:11px;color:var(--tx);min-width:64px;text-align:right">Cap</span>
          </div>
          <button class="bok" style="width:100%" onclick="saveContribution()">Save Contribution %</button>

          <div class="setDiv"></div>
          <label class="mlbl">Stock Selection Limits</label>
          <div class="setHint">Applies to the stocks the AI picks for you. Leave a box empty (or 0) for no limit. <b>Price range</b> skips stocks whose current price is outside it; <b>units</b> caps the shares in one trade (and, if you set a minimum, skips trades that can't reach it).</div>
          <div style="display:flex;gap:6px;margin-bottom:8px">
            <input class="minp" id="selPriceMin" type="number" min="0" step="any" placeholder="Min price &#8377;" style="height:34px"/>
            <input class="minp" id="selPriceMax" type="number" min="0" step="any" placeholder="Max price &#8377;" style="height:34px"/>
          </div>
          <div style="display:flex;gap:6px;margin-bottom:12px">
            <input class="minp" id="selUnitsMin" type="number" min="0" step="1" placeholder="Min units" style="height:34px"/>
            <input class="minp" id="selUnitsMax" type="number" min="0" step="1" placeholder="Max units" style="height:34px"/>
          </div>
          <button class="bok" style="width:100%" onclick="saveSelectionLimits()">Save Selection Limits</button>

          <div class="setDiv"></div>
          <label class="mlbl">Weekend &amp; Holiday Trading</label>
          <label class="authChk" style="margin-bottom:6px"><input type="checkbox" id="ntdOn" onchange="saveNonTradingDay()"/> Trade on weekends &amp; market holidays</label>
          <div class="setHint">When ON, every non-trading-day restriction is bypassed for your account on a Saturday, Sunday or listed holiday: the AI auto-trader runs, pending orders keep filling, and the 3:10 / 3:30 PM cutoffs and market-hours checks do not apply &mdash; so special sessions such as <b>Diwali Muhurat</b> work. Normal trading days are unaffected, so it is safe to leave on, but prices on a closed market are not live: switch it off when the session ends.</div>

          <div class="setDiv"></div>
          <label class="mlbl">Live Auto-trading (real money)</label>
          <label class="authChk" style="margin-bottom:6px"><input type="checkbox" id="liveAutoOn" onchange="saveLiveAuto()"/> Let the AI place real orders on my Zerodha account</label>
          <div class="setHint"><b>OFF (default):</b> the AI never touches your Live account &mdash; you can still view the Live tab and place live orders yourself. <b>ON:</b> when Trading Account is <b>Live</b> or <b>Both</b> and your execution mode is <b>Auto</b>, the AI buys/averages and exits (Action = SELL, profit-lock, pre-close, square-off) on your real account. It needs a connected broker and relay, and uses your per-trade cap and contribution %. Switch it off here any time.</div>

          <div style="border-top:1px solid var(--bd);margin:16px 0 12px"></div>
          <label class="mlbl">Target Hold Period</label>
          <div style="font-size:10px;color:var(--dm);margin-bottom:10px;line-height:1.5">Advisory only — shared with the AI as a guideline alongside price action and news, never a rigid rule that forces an exit on its own.</div>
          <label class="authChk" style="margin-bottom:8px"><input type="checkbox" id="holdShortOn"/> Use for Short-term</label>
          <div style="display:flex;align-items:center;gap:10px;margin-bottom:14px">
            <input type="range" id="holdShortWeeks" min="1" max="8" step="1" value="2" style="flex:1" oninput="document.getElementById('holdShortLbl').textContent=this.value+'wk'"/>
            <span id="holdShortLbl" style="font-size:11px;color:var(--tx);min-width:34px;text-align:right">2wk</span>
          </div>
          <label class="authChk" style="margin-bottom:8px"><input type="checkbox" id="holdLongOn"/> Use for Long-term</label>
          <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px">
            <input type="range" id="holdLongWeeks" min="1" max="52" step="1" value="16" style="flex:1" oninput="document.getElementById('holdLongLbl').textContent=fmtHoldWeeks_(this.value)"/>
            <span id="holdLongLbl" style="font-size:11px;color:var(--tx);min-width:60px;text-align:right">16wk (~4mo)</span>
          </div>
          <button class="bok" style="width:100%" onclick="saveHoldPeriod()">Save Hold Period</button>

          <div style="border-top:1px solid var(--bd);margin:16px 0 12px"></div>
          <label class="mlbl">Trading Account</label>
          <div style="font-size:10px;color:var(--dm);margin-bottom:10px;line-height:1.5">Mock is paper trading &mdash; safe to explore, never touches real money. Live places real orders through your connected broker. Choose which shows up in the Portfolio tab; pick Both to get a swap control there. <b>This also decides where the AI\'s auto-trades go:</b> Mock only &rarr; mock orders only; Live only &rarr; real broker orders only; Both &rarr; both &mdash; real orders are placed only if <b>Live auto-trading</b> (below) is switched on and your broker is connected.</div>
          <div class="mtyps" id="tradingAcctBtns" style="margin-bottom:4px">
            <button class="tybtn" data-acct="mock" onclick="setTradingAccountPref('mock')">Mock only</button>
            <button class="tybtn" data-acct="live" onclick="setTradingAccountPref('live')">Live only</button>
            <button class="tybtn" data-acct="both" onclick="setTradingAccountPref('both')">Both</button>
          </div>
        </div>
      </details>

      <details class="setSec" id="brkSec">
        <summary>Brokerage</summary>
        <div class="setBody">
          <label class="mlbl">Broker</label>
          <input class="minp" id="brkName" type="text" maxlength="40" style="height:34px" placeholder="e.g. Zerodha"/>
          <label class="authChk" style="margin:10px 0 6px"><input type="checkbox" id="brkApply"/> Use these charges in portfolio P&amp;L (realized &amp; unrealized)</label>
          <div class="setHint">When on, Holdings, Positions and the Trade Log show profit <b>after</b> brokerage, STT/CTT, exchange, SEBI, stamp duty and GST. Open positions also allow for the estimated cost of exiting. Rates are saved per user &mdash; edit them to match your broker.</div>
          <label class="mlbl">Exchange used for transaction charges</label>
          <select class="minp" id="brkExch" style="height:34px;margin-bottom:6px"><option value="NSE">NSE</option><option value="BSE">BSE</option></select>
          <div id="brkGrid"></div>
          <div class="setHint">Zerodha's schedule: delivery has zero brokerage; intraday is 0.03% or &#8377;20 per executed order, whichever is lower. GST applies to brokerage + SEBI + transaction charges. (&#8377;10 per crore = 0.0001%.) <b>IPFT</b> (NSE Investor Protection Fund Trust) is charged on equity &amp; futures turnover, with GST on top; it defaults to the figure you supplied, &#8377;0.01 per crore &mdash; NSE has notified &#8377;10 per crore since 1 Apr 2023, so change it if your broker bills that. IPFT is NSE-only and is skipped when the exchange above is BSE.</div>

          <div class="setDiv"></div>
          <label class="mlbl">Margin Privilege (intraday leverage)</label>
          <label class="authChk" style="margin-bottom:8px"><input type="checkbox" id="brkLevOn"/> Enable margin privilege on intraday trades</label>
          <div style="display:flex;align-items:center;gap:10px;margin-bottom:6px">
            <input type="range" id="brkLev" min="1" max="5" step="0.5" value="5" style="flex:1" oninput="updBrkLev_()"/>
            <span id="brkLevLbl" style="font-size:11px;color:var(--tx);min-width:34px;text-align:right">5x</span>
          </div>
          <div class="setHint" id="brkLevHint"></div>
          <div style="display:flex;gap:6px">
            <button class="bcx" style="flex-shrink:0" onclick="resetBrokerageDefaults()">Zerodha defaults</button>
            <button class="bok" style="flex:1" onclick="saveBrokerage()">Save Brokerage</button>
          </div>
        </div>
      </details>

      <details class="setSec" id="brokerSec">
        <summary>Broker (Live)</summary>
        <div class="setBody">
          <div class="setHint">Connect <b>your own</b> Zerodha Kite Connect app so the Live tab trades your account. Kite only accepts API orders from a static IP registered on that app, so you bring your own order relay running on your own static IP (see ARCHITECTURE_LIVE_TRADING.md). Other users' details never touch yours.</div>
          <button class="bok" style="width:100%;margin-bottom:10px" onclick="openBrokerGuide()">&#128216; Setup guide &mdash; how to get each detail</button>
          <div id="brokerStatusLine" class="setHint" style="color:var(--tx)">Loading&hellip;</div>
          <label class="mlbl">Kite Connect API key</label>
          <input class="minp" id="bkApiKey" type="text" autocomplete="off" style="height:34px;margin-bottom:6px" placeholder="API key"/>
          <label class="mlbl">API secret</label>
          <input class="minp" id="bkApiSecret" type="password" autocomplete="new-password" style="height:34px;margin-bottom:6px" placeholder="API secret"/>
          <label class="mlbl">Zerodha user ID</label>
          <input class="minp" id="bkUserId" type="text" autocomplete="off" maxlength="12" style="height:34px;margin-bottom:6px" placeholder="e.g. AB1234"/>
          <label class="mlbl">Order relay URL (your static-IP server)</label>
          <input class="minp" id="bkRelayUrl" type="text" autocomplete="off" style="height:34px;margin-bottom:6px" placeholder="https://your-relay.example.com/place-order"/>
          <label class="mlbl">Relay secret</label>
          <input class="minp" id="bkRelaySecret" type="password" autocomplete="new-password" style="height:34px;margin-bottom:10px" placeholder="Shared secret your relay expects"/>
          <label class="mlbl">Your own static-IP proxy URL (optional)</label>
          <input class="minp" id="bkProxy" type="password" autocomplete="new-password" style="height:34px;margin-bottom:6px" placeholder="http://USER:PASS@host:port — leave blank to use the relay's own static IP"/>
          <div class="setHint">Leave blank if you use the dashboard owner's relay <i>and</i> static IP (family). If you have your own TrueIP (or similar) address, paste its proxy URL here: the relay then sends <b>your</b> orders from <b>your</b> IP, which is the one to whitelist in your Zerodha developer account.</div>
          <div style="display:flex;gap:6px;margin-bottom:6px"><button class="bcx" onclick="testBrokerProxy()">Test &mdash; show my IP</button><button class="bcx" onclick="clearBrokerProxy()">Remove my proxy</button></div>
          <div id="bkProxyMsg" class="setHint" style="margin-bottom:10px"></div>

          <div class="setDiv"></div>
          <label class="mlbl">Automatic daily login (optional)</label>
          <div class="setWarn">Kite access tokens expire every morning, so a hands-free login means this app logs in to Zerodha for you using your saved user ID, password and TOTP secret. <b>Zerodha does not provide or endorse automated login</b> and can change or block it at any time; if it fails you can still press <i>Connect Zerodha</i> in the Live tab. Your password and TOTP secret are stored encrypted, but anyone who can open this Apps Script project can in principle decrypt them &mdash; only enable this if you accept that risk.</div>
          <label class="authChk" style="margin-bottom:8px"><input type="checkbox" id="bkAck"/> I understand and accept this</label>
          <label class="mlbl">Zerodha password</label>
          <input class="minp" id="bkPwd" type="password" autocomplete="new-password" style="height:34px;margin-bottom:6px" placeholder="Zerodha password"/>
          <div class="setHint"><b>Not needed to connect normally.</b> Pressing <i>Connect Zerodha</i> each morning works without a password or TOTP here. These two fields are only for the hands-free login.</div>
          <label class="mlbl">Zerodha TOTP secret (optional &mdash; only for automatic login)</label>
          <input class="minp" id="bkTotp" type="password" autocomplete="new-password" style="height:34px;margin-bottom:8px" placeholder="TOTP secret key"/>
          <label class="authChk" style="margin-bottom:10px"><input type="checkbox" id="bkAuto"/> Connect automatically every trading morning</label>
          <button class="bok" style="width:100%;margin-bottom:6px" onclick="saveBrokerSettings()">Save Broker Settings</button>
          <div style="display:flex;gap:6px;flex-wrap:wrap">
            <button class="bcx" onclick="brokerConnectNow()">Connect now</button>
            <button class="bcx" onclick="clearBrokerAuto()">Remove saved login</button>
            <button class="bcx" style="color:var(--rd);border-color:var(--rd)" onclick="removeBroker()">Remove broker</button>
          </div>
        </div>
      </details>

      <details class="setSec">
        <summary>Referral</summary>
        <div class="setBody">
          <label class="mlbl">Your Referral Link <span style="color:var(--dm)" id="refCount"></span></label>
          <div style="display:flex;gap:6px">
            <input class="minp" id="refLinkFull" type="text" style="height:34px;font-size:10px" readonly/>
            <button class="bok" style="padding:0 14px;flex-shrink:0" onclick="copyReferralLink()">Copy</button>
          </div>
        </div>
      </details>

      <details class="setSec">
        <summary>Security &amp; Account</summary>
        <div class="setBody">
          <label class="mlbl">Two-step verification (MFA)</label>
          <div id="mfaStatusLine" class="setHint" style="color:var(--tx)">Checking&hellip;</div>
          <div id="mfaOffBox" style="display:none;margin-bottom:14px">
            <div class="setHint">Adds a 6-digit code from an authenticator app (Google Authenticator, Authy, Microsoft Authenticator, 1Password&hellip;) to every login.</div>
            <input class="minp" id="mfaPwdOn" type="password" placeholder="Current password" autocomplete="current-password" style="height:34px;margin-bottom:6px"/>
            <button class="bok" style="width:100%" onclick="mfaStart()">Turn on two-step verification</button>
          </div>
          <div id="mfaSetupBox" style="display:none;margin-bottom:14px">
            <div class="setHint"><b>1.</b> In your authenticator app choose <i>Add account</i> and scan this code (or type the key). <b>2.</b> Enter the 6-digit code it shows.</div>
            <div id="mfaQr" style="background:#fff;display:inline-block;padding:18px;border-radius:8px;margin-bottom:8px;line-height:0"></div>
            <div class="setHint"><b>Google Authenticator:</b> + &rarr; Scan a QR code. <b>Microsoft Authenticator:</b> + &rarr; Other account (Google, Facebook, etc.) &rarr; scan, or choose <i>Or enter code manually</i> and paste the key below (Account name: anything, Secret key: the key). Hold the phone about 20&nbsp;cm from the screen.</div>
            <div style="display:flex;gap:6px;margin-bottom:6px"><input class="minp" id="mfaSecretTxt" type="text" readonly style="height:34px;font-size:11px;margin:0;flex:1" onclick="this.select()"/><button class="bcx" onclick="mfaCopyKey(this)">Copy key</button></div>
            <input class="minp" id="mfaConfirmCode" type="text" inputmode="numeric" maxlength="7" placeholder="6-digit code" autocomplete="one-time-code" style="height:34px;margin-bottom:6px"/>
            <button class="bok" style="width:100%" onclick="mfaConfirm()">Verify &amp; turn on</button>
          </div>
          <div id="mfaCodesBox" style="display:none;margin-bottom:14px">
            <div class="setWarn"><b>Save these recovery codes now.</b> Each works once if you lose your phone. They are not shown again.</div>
            <pre id="mfaCodesPre" style="background:var(--sf2);border:1px solid var(--bd);border-radius:6px;padding:10px;font-size:12px;letter-spacing:1px;margin:0 0 8px"></pre>
            <div style="display:flex;gap:6px"><button class="bcx" onclick="mfaCopyCodes()">Copy</button><button class="bok" style="flex:1" onclick="mfaCodesDone()">I've saved them</button></div>
          </div>
          <div id="mfaOnBox" style="display:none;margin-bottom:14px">
            <input class="minp" id="mfaPwdOff" type="password" placeholder="Current password" autocomplete="current-password" style="height:34px;margin-bottom:6px"/>
            <input class="minp" id="mfaCodeOff" type="text" placeholder="Authenticator code (or recovery code)" autocomplete="one-time-code" style="height:34px;margin-bottom:6px"/>
            <div style="display:flex;gap:6px"><button class="bcx" style="flex:1" onclick="mfaNewCodes()">New recovery codes</button><button class="bcx" style="flex:1;color:var(--rd);border-color:var(--rd)" onclick="mfaTurnOff()">Turn off</button></div>
          </div>
          <div class="setDiv" style="margin-top:4px"></div>
          <label class="mlbl">Change Password</label>
          <input class="minp" id="profCurPass" type="password" placeholder="Current password" style="height:34px;margin-bottom:6px"/>
          <input class="minp" id="profNewPass" type="password" placeholder="New password" style="height:34px;margin-bottom:6px"/>
          <input class="minp" id="profNewPass2" type="password" placeholder="Confirm new password" style="height:34px"/>
          <button class="bok" style="margin:8px 0 14px;width:100%" onclick="saveProfilePassword()">Change Password</button>
          <div class="mft" style="justify-content:space-between">
            <button class="bcx" onclick="doLogout()">Log Out</button>
            <button class="bcx" style="color:var(--rd);border-color:var(--rd)" onclick="confirmDeleteAccount()">Delete Account</button>
          </div>
        </div>
      </details>
    </div>
  </div>
</div>

<div class="mover" id="trafficM">
  <div class="modal" style="max-width:560px">
    <div class="mhd"><span class="mttl">Traffic Stats</span><button class="mx" onclick="closeTraffic()">&#10005;</button></div>
    <div class="mbdy">
      <div id="trafficSummary" style="display:flex;gap:18px;margin-bottom:14px;flex-wrap:wrap">
        <div><div style="font-size:20px;font-weight:700" id="trafTotal">-</div><div style="font-size:10px;color:var(--dm);text-transform:uppercase;letter-spacing:1px">Total Visits (all time)</div></div>
        <div><div style="font-size:20px;font-weight:700" id="trafUnique">-</div><div style="font-size:10px;color:var(--dm);text-transform:uppercase;letter-spacing:1px">Unique Users</div></div>
        <div><div style="font-size:20px;font-weight:700" id="trafToday">-</div><div style="font-size:10px;color:var(--dm);text-transform:uppercase;letter-spacing:1px">Today</div></div>
        <div><div style="font-size:20px;font-weight:700" id="trafRange">-</div><div style="font-size:10px;color:var(--dm);text-transform:uppercase;letter-spacing:1px" id="trafRangeLbl">In Range</div></div>
      </div>
      <div class="mtyps" id="trafRangeBtns" style="margin-bottom:12px">
        <button class="tybtn" data-r="7" onclick="setTrafficRange(7)">7D</button>
        <button class="tybtn" data-r="30" onclick="setTrafficRange(30)">30D</button>
        <button class="tybtn" data-r="90" onclick="setTrafficRange(90)">90D</button>
        <button class="tybtn" data-r="0" onclick="setTrafficRange(0)">All Time</button>
      </div>
      <div id="trafChart"></div>
    </div>
  </div>
</div>

<div id="loader"><div class="llo">Signal Performance</div><div class="lbar"><div class="lfil"></div></div><div class="ltxt">Loading...</div></div>

<div class="shell">
<header>
  <div class="hl">
    <button class="mob-ham" onclick="toggleSB()">&#9776;</button>
    <div class="lgrp"><span class="logo">Signal Performance</span><span class="slo">NSE / BSE &middot; Real-Time</span></div>
    <div class="tnav">
      <button class="tbtn active" data-tab="home"     onclick="switchTopTab('home')">&#8962; Home</button>
      <button class="tbtn" data-tab="intraday"        onclick="switchTopTab('intraday')">&#9889; Intraday</button>
      <button class="tbtn" data-tab="shortterm"       onclick="switchTopTab('shortterm')">&#8645; Short-term</button>
      <button class="tbtn" data-tab="longterm"        onclick="switchTopTab('longterm')">&#9651; Long-term</button>
      <button class="tbtn" data-tab="watchlist"        onclick="switchTopTab('watchlist')">&#9733; Watchlist</button>
      <button class="tbtn" data-tab="portfolio"        onclick="switchTopTab('portfolio')">&#128188; Portfolio</button>
    </div>
  </div>
  <div class="hsp"></div>
  <div class="tcks">
    <div class="htck"><span class="htck-n">SENSEX</span>&nbsp;<span class="htck-p" id="tSX">-</span>&nbsp;<span class="htck-c" id="tSXc"></span></div>
    <div class="htck-sep"></div>
    <div class="htck"><span class="htck-n">NIFTY</span>&nbsp;<span class="htck-p" id="tNF">-</span>&nbsp;<span class="htck-c" id="tNFc"></span></div>
    <div class="htck-sep"></div>
    <div class="htck"><span class="htck-n">BANKNIFTY</span>&nbsp;<span class="htck-p" id="tBN">-</span>&nbsp;<span class="htck-c" id="tBNc"></span></div>
  </div>
  <div class="hr">
    <button class="mob-ai" onclick="toggleAI()">&#10024;</button>
    <span class="mkt-tag" id="ntdBanner" style="display:none" title="Weekend &amp; holiday trading is ON (Settings → Trading Preferences)">Special-session trading ON</span>
    <div class="livb"><span class="livd"></span><span id="livLbl">Live</span></div>
    <div class="clk" id="clk">--:--:--</div>
    <button class="icb" onclick="toggleTheme()"><span id="thI">&#9681;</span></button>
  </div>
</header>

<div class="pbar">
  <span class="plbl">Period</span>
  <div class="mgrp" id="mgrp">
    <button class="mbtn active" data-m="DAYS">Days</button>
    <button class="mbtn" data-m="WEEKS">Weeks</button>
    <button class="mbtn" data-m="MONTHS">Months</button>
    <button class="mbtn" data-m="YEARS">Years</button>
  </div>
  <input class="pinp" id="pval" type="number" value="1" min="1" max="999"/>
  <div class="bsp"></div>
  <span class="srng" id="srng">Today</span>
  <div class="bsp" id="dnBsp" style="display:none"></div>
  <div class="dnav-grp" id="dnavGrp">
    <span class="plbl">Date</span>
    <button class="dnb" onclick="shiftDate(-1)">&#9664;</button>
    <input class="datei" type="date" id="gDate" onchange="onDateChange(this.value)"/>
    <button class="dnb tod" onclick="gotoToday()">&#9679;</button>
    <button class="dnb" onclick="shiftDate(1)">&#9654;</button>
    <span class="plbl" style="margin-left:4px">At</span>
    <input class="timei" id="asOfTime" type="time" onchange="onTimeChange(this.value)"/>
  </div>
  <div class="bsp"></div>
  <span class="plbl">Price</span>
  <div class="ptog">
    <button class="ptb active" data-pm="ltp"  onclick="setPM('ltp')">LTP</button>
    <button class="ptb"        data-pm="high" onclick="setPM('high')">High</button>
  </div>
  <span class="mkt-tag" id="mktTag" style="display:none">Closed</span>
</div>

<div class="krow">
  <div class="kpi" style="--ka:var(--bl)" data-ki="0"><div class="kpiTxt"><div class="klbl" id="k0l">Total</div><div class="kval" id="k0">-</div><div class="ksub" id="k0s"></div></div><div class="kring"><svg viewBox="0 0 36 36"><circle class="kringBg" cx="18" cy="18" r="15"/><circle class="kringFg" id="k0rc" cx="18" cy="18" r="15"/></svg><span class="kringPct" id="k0rp">0%</span></div></div>
  <div class="kpi" style="--ka:var(--gr)" data-ki="1"><div class="kpiTxt"><div class="klbl" id="k1l">In Profit</div><div class="kval" id="k1">-</div><div class="ksub" id="k1s"></div></div><div class="kring"><svg viewBox="0 0 36 36"><circle class="kringBg" cx="18" cy="18" r="15"/><circle class="kringFg" id="k1rc" cx="18" cy="18" r="15"/></svg><span class="kringPct" id="k1rp">0%</span></div></div>
  <div class="kpi" style="--ka:var(--rd)" data-ki="2"><div class="kpiTxt"><div class="klbl" id="k2l">In Loss</div><div class="kval" id="k2">-</div><div class="ksub" id="k2s"></div></div><div class="kring"><svg viewBox="0 0 36 36"><circle class="kringBg" cx="18" cy="18" r="15"/><circle class="kringFg" id="k2rc" cx="18" cy="18" r="15"/></svg><span class="kringPct" id="k2rp">0%</span></div></div>
  <div class="kpi" style="--ka:var(--gr)" data-ki="3"><div class="kpiTxt"><div class="klbl" id="k3l">Target Hits</div><div class="kval" id="k3">-</div><div class="ksub" id="k3s"></div></div><div class="kring"><svg viewBox="0 0 36 36"><circle class="kringBg" cx="18" cy="18" r="15"/><circle class="kringFg" id="k3rc" cx="18" cy="18" r="15"/></svg><span class="kringPct" id="k3rp">0%</span></div></div>
  <div class="kpi" style="--ka:var(--rd)" data-ki="4"><div class="kpiTxt"><div class="klbl" id="k4l">SL Hits</div><div class="kval" id="k4">-</div><div class="ksub" id="k4s"></div></div><div class="kring"><svg viewBox="0 0 36 36"><circle class="kringBg" cx="18" cy="18" r="15"/><circle class="kringFg" id="k4rc" cx="18" cy="18" r="15"/></svg><span class="kringPct" id="k4rp">0%</span></div></div>
  <div class="kpi" style="--ka:var(--pu)" data-ki="5"><div class="kpiTxt"><div class="klbl" id="k5l">Exited</div><div class="kval" id="k5">-</div><div class="ksub" id="k5s"></div></div><div class="kring"><svg viewBox="0 0 36 36"><circle class="kringBg" cx="18" cy="18" r="15"/><circle class="kringFg" id="k5rc" cx="18" cy="18" r="15"/></svg><span class="kringPct" id="k5rp">0%</span></div></div>
</div>

<div id="dr"><div class="drhd"><div style="display:flex;align-items:center;gap:10px"><span class="drttl" id="drttl">-</span><span class="drcnt" id="drcnt"></span></div><button class="drx" onclick="closeDrST()">&#10005;</button></div><input class="drsc" id="drsc" placeholder="Search, or status:exited days&gt;10 return&lt;-2 ..."/><div class="drbdy"><table><thead><tr id="drThead"><th>#</th><th data-col="symbol" onclick="sortDr('symbol')">Symbol</th><th data-col="since" onclick="sortDr('since')">Since</th><th data-col="entry" onclick="sortDr('entry')">Entry</th><th data-col="ltp" onclick="sortDr('ltp')">LTP</th><th data-col="sl" onclick="sortDr('sl')">SL</th><th data-col="target" onclick="sortDr('target')">Target</th><th data-col="high" onclick="sortDr('high')">High</th><th data-col="low" onclick="sortDr('low')">Low</th><th data-col="return" onclick="sortDr('return')">Return</th><th data-col="thit" onclick="sortDr('thit')">T.Hit</th><th data-col="slhit" onclick="sortDr('slhit')">SL.Hit</th><th data-col="days" onclick="sortDr('days')">Days</th><th data-col="status" onclick="sortDr('status')">Status</th><th data-col="exited" onclick="sortDr('exited')">Exited</th></tr></thead><tbody id="drbdy"></tbody></table></div></div>
<div id="idr"><div class="drhd"><div style="display:flex;align-items:center;gap:10px"><span class="drttl" id="idrttl">-</span><span class="drcnt" id="idrcnt"></span></div><button class="drx" onclick="closeDrI()">&#10005;</button></div><input class="drsc" id="idrsc" placeholder="Search, or status:exited days&gt;30 return&lt;-2 ..."/><div class="drbdy"><table><thead><tr id="idrThead"><th>#</th><th data-col="symbol" onclick="sortIDr('symbol')">Symbol</th><th data-col="order" onclick="sortIDr('order')">Order</th><th data-col="since" onclick="sortIDr('since')">Date &amp; Time</th><th data-col="entry" onclick="sortIDr('entry')">Entry</th><th data-col="ltp" onclick="sortIDr('ltp')">LTP</th><th data-col="sl" onclick="sortIDr('sl')">SL</th><th data-col="high" onclick="sortIDr('high')">High</th><th data-col="low" onclick="sortIDr('low')">Low</th><th data-col="return" onclick="sortIDr('return')">Return</th><th data-col="target" onclick="sortIDr('target')">Target</th><th data-col="thit" onclick="sortIDr('thit')">T.Hit</th><th data-col="slhit" onclick="sortIDr('slhit')">SL.Hit</th><th data-col="days" onclick="sortIDr('days')">Mins</th><th data-col="status" onclick="sortIDr('status')">Status</th></tr></thead><tbody id="idrbdy"></tbody></table></div></div>
<div id="ltdr"><div class="drhd"><div style="display:flex;align-items:center;gap:10px"><span class="drttl" id="ltdrttl">-</span><span class="drcnt" id="ltdrcnt"></span></div><button class="drx" onclick="closeDrLT()">&#10005;</button></div><input class="drsc" id="ltdrsc" placeholder="Search, or status:exited days&gt;10 return&lt;-2 ..."/><div class="drbdy"><table><thead><tr id="ltdrThead"><th>#</th><th data-col="symbol" onclick="sortLtDr('symbol')">Symbol</th><th data-col="since" onclick="sortLtDr('since')">Since</th><th data-col="entry" onclick="sortLtDr('entry')">Entry</th><th data-col="ltp" onclick="sortLtDr('ltp')">LTP</th><th data-col="sl" onclick="sortLtDr('sl')">SL</th><th data-col="target" onclick="sortLtDr('target')">Target</th><th data-col="high" onclick="sortLtDr('high')">High</th><th data-col="low" onclick="sortLtDr('low')">Low</th><th data-col="return" onclick="sortLtDr('return')">Return</th><th data-col="thit" onclick="sortLtDr('thit')">T.Hit</th><th data-col="slhit" onclick="sortLtDr('slhit')">SL.Hit</th><th data-col="days" onclick="sortLtDr('days')">Days</th><th data-col="status" onclick="sortLtDr('status')">Status</th><th data-col="exited" onclick="sortLtDr('exited')">Exited</th></tr></thead><tbody id="ltdrbdy"></tbody></table></div></div>

<div id="cnt-home" style="display:none">
  <div class="php"><div class="phico">&#9711;</div><div class="phttl" style="color:#4da6ff">Home Dashboard</div>
    <div class="phsub">Under Construction</div>
    <div class="phch">
      <span class="phc" onclick="switchTopTab('intraday')">&#9889; Intraday</span>
      <span class="phc" onclick="switchTopTab('shortterm')">&#8645; Short-term</span>
      <span class="phc" onclick="switchTopTab('longterm')">&#9651; Long-term</span>
    </div>
  </div>
</div>

<div id="cnt-intraday" style="display:none">
  <div class="mgrid">
    <div class="pan"><div class="phd"><span class="pttl">Best P&amp;L</span><span class="bdg bc">&#8593; Top</span></div><table><thead><tr><th>#</th><th>Symbol</th><th>Order</th><th>Entry</th><th>LTP</th><th>Return</th><th>Status</th></tr></thead><tbody id="itopBdy"><tr class="emr"><td colspan="7">Loading...</td></tr></tbody></table></div>
    <div class="pan"><div class="phd"><span class="pttl">Worst P&amp;L</span><span class="bdg br">&#8595; Loss</span></div><table><thead><tr><th>#</th><th>Symbol</th><th>Order</th><th>Entry</th><th>LTP</th><th>Return</th><th>Status</th></tr></thead><tbody id="iwstBdy"><tr class="emr"><td colspan="7">Loading...</td></tr></tbody></table></div>
    <div class="ccol">
      <div class="pan"><div class="phd"><span class="pttl">Session Ratio</span><span class="bdg bc" id="isrBdg">-</span></div>
        <div class="dwrap"><svg viewBox="0 0 100 100"><circle cx="50" cy="50" r="38" fill="none" stroke-width="12" style="stroke:var(--bd)"/><circle id="iarc" cx="50" cy="50" r="38" fill="none" stroke="#00d4ff" stroke-width="12" stroke-dasharray="0 238.76" stroke-linecap="round" style="transition:stroke-dasharray 1.2s cubic-bezier(.16,1,.3,1)"/></svg><div class="dctr"><span class="dpct" id="idpct" style="color:var(--cy)">0%</span><span class="dlbl">Profit</span></div></div>
        <div class="legi"><div class="li"><div class="ld" style="background:var(--gr)"></div><span class="lt">In Profit</span><span class="lv" id="ilgP" style="color:var(--gr)">-</span></div><div class="li"><div class="ld" style="background:var(--rd)"></div><span class="lt">In Loss</span><span class="lv" id="ilgL" style="color:var(--rd)">-</span></div><div class="li"><div class="ld" style="background:var(--gr);opacity:.6"></div><span class="lt">BUY</span><span class="lv" id="ilgB" style="color:var(--gr)">-</span></div><div class="li"><div class="ld" style="background:var(--rd);opacity:.6"></div><span class="lt">Short-sell</span><span class="lv" id="ilgS" style="color:var(--rd)">-</span></div></div></div>
      <div class="pan"><div class="phd"><span class="pttl">Session Stats</span><span class="bdg bc">Intraday</span></div><div class="bsec" id="ibsec"></div></div>
    </div>
  </div>
</div>

<div id="cnt-active" style="display:none">
  <div class="mgrid">
<!-- ST: LTP moved next to Entry per request -->
    <div class="pan"><div class="phd"><span class="pttl">Top Performers</span><span class="bdg bg">&#8593; Best</span></div><table><thead><tr><th>#</th><th>Symbol</th><th>Entry</th><th>LTP</th><th>SL</th><th>High</th><th>Return</th><th>Status</th></tr></thead><tbody id="topBdy"><tr class="emr"><td colspan="8">Loading...</td></tr></tbody></table></div>
    <div class="pan"><div class="phd"><span class="pttl">Worst Performers</span><span class="bdg br">&#8595; Worst</span></div><table><thead><tr><th>#</th><th>Symbol</th><th>Entry</th><th>LTP</th><th>SL</th><th>High</th><th>Return</th><th>Status</th></tr></thead><tbody id="wstBdy"><tr class="emr"><td colspan="8">Loading...</td></tr></tbody></table></div>
    <div class="ccol">
      <div class="pan"><div class="phd"><span class="pttl">Profit Ratio</span><span class="bdg bg" id="srBdg">-</span></div>
        <div class="dwrap"><svg viewBox="0 0 100 100"><circle cx="50" cy="50" r="38" fill="none" stroke-width="12" style="stroke:var(--bd)"/><circle id="darc" cx="50" cy="50" r="38" fill="none" stroke="#00e68a" stroke-width="12" stroke-dasharray="0 238.76" stroke-linecap="round" style="transition:stroke-dasharray 1.2s cubic-bezier(.16,1,.3,1)"/></svg><div class="dctr"><span class="dpct" id="dpct">0%</span><span class="dlbl">Profit</span></div></div>
        <div class="legi"><div class="li"><div class="ld" style="background:var(--gr)"></div><span class="lt">In Profit</span><span class="lv" id="lgP" style="color:var(--gr)">-</span></div><div class="li"><div class="ld" style="background:var(--rd)"></div><span class="lt">In Loss</span><span class="lv" id="lgL" style="color:var(--rd)">-</span></div><div class="li"><div class="ld" style="background:var(--am)"></div><span class="lt">Running</span><span class="lv" id="lgR" style="color:var(--am)">-</span></div><div class="li"><div class="ld" style="background:var(--pu)"></div><span class="lt">Exited</span><span class="lv" id="lgE" style="color:var(--pu)">-</span></div></div></div>
      <div class="pan"><div class="phd"><span class="pttl">Stats</span><span class="bdg ba">Breakdown</span></div><div class="bsec" id="bsec"></div></div>
    </div>
  </div>
</div>

<div id="cnt-longterm" style="display:none">
  <div class="mgrid">
    <div class="pan"><div class="phd"><span class="pttl">Top Holdings</span><span class="bdg ba">&#8593; Best</span></div><table><thead><tr><th>#</th><th>Symbol</th><th>Entry</th><th>LTP</th><th>High</th><th>Return</th><th>Status</th></tr></thead><tbody id="lttopBdy"><tr class="emr"><td colspan="7">Loading...</td></tr></tbody></table></div>
    <div class="pan"><div class="phd"><span class="pttl">Worst Holdings</span><span class="bdg br">&#8595; Worst</span></div><table><thead><tr><th>#</th><th>Symbol</th><th>Entry</th><th>LTP</th><th>High</th><th>Return</th><th>Status</th></tr></thead><tbody id="ltwstBdy"><tr class="emr"><td colspan="7">-</td></tr></tbody></table></div>
    <div class="ccol">
      <div class="pan"><div class="phd"><span class="pttl">Holdings Ratio</span><span class="bdg ba" id="ltSrBdg">-</span></div>
        <div class="dwrap"><svg viewBox="0 0 100 100"><circle cx="50" cy="50" r="38" fill="none" stroke-width="12" style="stroke:var(--bd)"/><circle id="ltarc" cx="50" cy="50" r="38" fill="none" stroke="var(--am)" stroke-width="12" stroke-dasharray="0 238.76" stroke-linecap="round" style="transition:stroke-dasharray 1.2s cubic-bezier(.16,1,.3,1)"/></svg><div class="dctr"><span class="dpct" id="ltpct" style="color:var(--am)">0%</span><span class="dlbl">Profit</span></div></div>
        <div class="legi"><div class="li"><div class="ld" style="background:var(--gr)"></div><span class="lt">Profitable</span><span class="lv" id="ltLP" style="color:var(--gr)">-</span></div><div class="li"><div class="ld" style="background:var(--rd)"></div><span class="lt">In Loss</span><span class="lv" id="ltLL" style="color:var(--rd)">-</span></div><div class="li"><div class="ld" style="background:var(--am)"></div><span class="lt">Running</span><span class="lv" id="ltLH" style="color:var(--am)">-</span></div><div class="li"><div class="ld" style="background:var(--pu)"></div><span class="lt">Exited</span><span class="lv" id="ltLE" style="color:var(--pu)">-</span></div></div></div>
      <div class="pan"><div class="phd"><span class="pttl">Holdings Stats</span><span class="bdg ba">Breakdown</span></div><div class="bsec" id="ltbsec"></div></div>
    </div>
  </div>
</div>

<div id="cnt-watchlist" style="display:none">
  <div class="pan">
    <div class="phd wlBar">
      <span class="pttl">Watchlist</span><span class="bdg ba" id="wlCount">0</span>
      <div class="wlAdd">
        <div class="acwrap"><input class="minp ac-input" id="wlSym" placeholder="Add symbol or company name..." autocomplete="off"/></div>
        <input class="minp wlNoteIn" id="wlNote" placeholder="Note (optional)" onkeydown="if(event.key==='Enter')addWLItem();"/>
        <button class="bok" onclick="addWLItem()">Add</button>
        <span id="wlAddStatus" class="wlSt"></span>
      </div>
      <input class="drsc wlSearch" id="wlSearch" placeholder="Search symbol or name..." oninput="filterWL()"/>
    </div>
    <div class="scrollTbl">
    <table><thead><tr><th>#</th><th>Symbol</th><th>Sector</th><th>Industry</th><th>LTP</th><th>Change</th><th>Change%</th><th>Open</th><th>High</th><th>Low</th><th>Prev Close</th><th>Volume (Cr)</th><th>Vol Avg (Cr)</th><th>52W High</th><th>52W Low</th><th>PE</th><th>EPS</th><th>Market Cap (Cr)</th></tr></thead>
      <tbody id="wlTblBdy"><tr class="emr"><td colspan="18">Loading...</td></tr></tbody>
    </table>
    </div>
  </div>

</div>



<div id="cnt-orderbook" style="display:none"></div>

<div id="cnt-portfolio" style="display:none">
  <!-- v3: ONE compact toolbar instead of five stacked bars (mode bar, add panel, margin strip, tab bar, section header) -->
  <div id="ptToolbar" class="ptTb">
    <div id="ptModeBar" class="ptModeBar">
      <span class="ptLbl">Mode</span>
      <button class="bcx" id="ptModeMockBtn" onclick="setTradingMode('mock')" style="border-color:var(--ac);color:var(--ac)">Mock</button>
      <button class="bcx" id="ptModeLiveBtn" onclick="setTradingMode('live')">Live</button>
      <span id="ptModeBanner" class="ptModeBanner" style="display:none">&#9888; Live &mdash; real money</span>
    </div>
    <div id="ptMarginStrip" class="mktStrip ptTbStrip"></div>
    <div id="liveToolbar" class="ptTbLive" style="display:none">
      <span class="ptLbl">Zerodha</span><span id="liveConnStatus" class="liveSt">Checking&hellip;</span>
      <button class="bok ptTbBtn" id="liveConnectBtn" onclick="connectBroker()" style="display:none">Connect Zerodha</button>
      <button class="bok ptTbBtn" id="liveSetupBtn" onclick="openBrokerSettings()" style="display:none">Set up your broker</button>
      <div id="liveMarginStrip" class="mktStrip ptTbStrip"></div>
      <div id="liveMarginsRow" style="display:none"></div>
    </div>
    <span class="ptSp"></span>
    <button class="bcx ptTbBtn" id="ptAddToggleBtn" onclick="toggleAddHolding()" title="Add a holding">&#43; Add</button>
    <button class="bcx ptIco" id="ptExportBtn" title="Download" aria-label="Download" onclick="exportPortfolioExcel()"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3v12m0 0l-4.5-4.5M12 15l4.5-4.5M4 20h16"/></svg></button>
  </div>

  <div id="liveTradingPanel" style="display:none">
    <div class="pan" id="liveOrderPan" style="display:none;margin-top:6px">
      <div class="phd"><span class="pttl">Place Live Order</span></div>
      <div style="padding:10px 13px;display:flex;gap:8px;flex-wrap:wrap;align-items:flex-end" id="liveOrderForm">
        <div style="width:170px"><label style="font-size:9px;color:var(--dm);display:block;margin-bottom:3px">Symbol</label><div class="acwrap" style="min-width:0"><input class="minp ac-input" id="liveOrdSymbol" placeholder="Symbol or company name..." autocomplete="off" style="margin-bottom:0"/></div></div>
        <div style="width:130px"><label style="font-size:9px;color:var(--dm);display:block;margin-bottom:3px">Category</label><select class="minp" id="liveOrdCategory" style="margin-bottom:0"><option value="intraday">Intraday (MIS)</option><option value="longterm">Delivery (CNC)</option></select></div>
        <div style="width:85px"><label style="font-size:9px;color:var(--dm);display:block;margin-bottom:3px">Side</label><select class="minp" id="liveOrdSide" style="margin-bottom:0"><option value="BUY">Buy</option><option value="SELL">Sell</option></select></div>
        <div style="width:75px"><label style="font-size:9px;color:var(--dm);display:block;margin-bottom:3px">Qty</label><input class="minp" id="liveOrdQty" type="number" min="1" style="margin-bottom:0"/></div>
        <div style="width:100px"><label style="font-size:9px;color:var(--dm);display:block;margin-bottom:3px">Type</label><select class="minp" id="liveOrdType" style="margin-bottom:0"><option value="MARKET">Market</option><option value="LIMIT">Limit</option></select></div>
        <div style="width:110px"><label style="font-size:9px;color:var(--dm);display:block;margin-bottom:3px">Price (limit only)</label><input class="minp" id="liveOrdPrice" type="number" step="0.01" style="margin-bottom:0"/></div>
        <button class="bok" id="liveOrdSubmitBtn" onclick="placeLiveOrder()" style="height:34px">Place Order</button>
      </div>
    </div>
    <div class="lptabs ptstrip" id="liveSubTabsBar">
      <button class="lptab ptsub-on" id="liveSubBtn-holdings" onclick="switchLiveSubtab('holdings')">Holdings <span class="bdg ba" id="liveHldCount">0</span></button>
      <button class="lptab" id="liveSubBtn-positions" onclick="switchLiveSubtab('positions')">Positions <span class="bdg ba" id="livePosCount">0</span></button>
      <button class="lptab" id="liveSubBtn-tradelog" onclick="switchLiveSubtab('tradelog')" title="Your live trades, grouped by stock">Trade Log <span class="bdg ba" id="liveTlCount">0</span></button>
      <span class="ptSp"></span>
      <span class="ptCtl" data-for="holdings"><input class="drsc" id="liveHldSearch" placeholder="Search symbol..." oninput="paintLiveSub_()"/></span>
      <span class="ptCtl" data-for="positions" style="display:none"><input class="drsc" id="livePosSearch" placeholder="Search symbol..." oninput="paintLiveSub_()"/><select class="drsc catSel" id="livePosFilter" onchange="paintLiveSub_()" title="Filter by product"><option value="">All</option><option value="MIS">Intraday (MIS)</option><option value="CNC">Delivery (CNC)</option></select></span>
      <span class="ptCtl" data-for="tradelog" style="display:none"><input class="drsc" id="liveTlSearch" placeholder="Search symbol or name..." oninput="paintLiveSub_()"/><select class="drsc catSel" id="liveTlCatFilter" onchange="paintLiveSub_()" title="Filter by timeframe"><option value="">All</option><option value="intraday">Intraday</option><option value="shortterm">Short-term</option><option value="longterm">Long-term</option></select></span>
    </div>
    <div class="pan ptBody" id="liveSub-holdings">
      <div id="liveHldSummary" class="ptSum"></div>
      <div class="scrollTbl"><table><thead><tr><th>#</th><th>Symbol</th><th>Qty</th><th>Avg Price</th><th>LTP</th><th>P&amp;L</th><th>Action</th></tr></thead><tbody id="liveHldTblBdy"><tr class="emr"><td colspan="7">Not connected.</td></tr></tbody></table></div>
    </div>
    <div class="pan ptBody" id="liveSub-positions" style="display:none">
      <div id="livePosSummary" class="ptSum"></div>
      <div class="scrollTbl"><table><thead><tr><th>#</th><th>Symbol</th><th>Product</th><th>Qty</th><th>Avg Price</th><th>LTP</th><th>P&amp;L</th><th>Action</th></tr></thead><tbody id="livePosTblBdy"><tr class="emr"><td colspan="8">Not connected.</td></tr></tbody></table></div>
    </div>
    <div class="pan ptBody" id="liveSub-tradelog" style="display:none">
      <div id="liveTlSummary" class="ptSum"></div>
      <div class="scrollTbl"><table><thead><tr><th>#</th><th>Symbol</th><th>Status</th><th>Category</th><th>Net Qty</th><th>Invested</th><th>Current/Exit Val.</th><th>P&amp;L</th><th>Return%</th><th>XIRR</th></tr></thead><tbody id="liveTlTblBdy"><tr class="emr"><td colspan="10">Not connected.</td></tr></tbody></table></div>
    </div>
  </div>

  <div class="pan" id="ptAddHoldingPan" style="display:none;margin-top:6px">
    <div class="addrow">
      <select class="minp" id="ptType" style="height:34px;max-width:110px" onchange="updPtEst()">
        <option value="BUY">Buy</option>
        <option value="SELL">Sell</option>
      </select>
      <select class="minp" id="ptCat" style="height:34px;max-width:120px">
        <option value="intraday">Intraday</option>
        <option value="shortterm">Short-term</option>
        <option value="longterm" selected>Long-term</option>
      </select>
      <div class="acwrap"><input class="minp ac-input" id="ptSym" placeholder="Symbol or company name..." autocomplete="off" style="height:34px"/></div>
      <input class="minp" id="ptQty" type="number" placeholder="Qty" style="height:34px;max-width:90px" oninput="updPtEst()"/>
      <input class="minp" id="ptPrice" type="number" placeholder="Price" style="height:34px;max-width:110px" oninput="updPtEst()"/>
      <input class="minp" id="ptDate" type="date" style="height:34px;max-width:140px"/>
      <button class="bok" onclick="addPTItem()">Add</button>
    </div>
    <div id="ptEst" style="padding:0 16px 8px;font-size:11px;color:var(--dm)"></div>
  </div>

  <div class="lptabs ptstrip" id="ptSubTabsBar">
    <button class="lptab ptsub-on" id="ptSubBtn-holdings" onclick="switchPortfolioSubtab('holdings')" title="Short-term &amp; long-term">Holdings <span class="bdg ba" id="ptCount">0</span></button>
    <button class="lptab" id="ptSubBtn-positions" onclick="switchPortfolioSubtab('positions')" title="Intraday and today's activity">Positions <span class="bdg ba" id="posCount">0</span></button>
    <button class="lptab" id="ptSubBtn-tradelog" onclick="switchPortfolioSubtab('tradelog')" title="Grouped by stock, including closed-out positions">Trade Log <span class="bdg ba" id="tlCount">0</span></button>
    <span class="ptSp"></span>
    <span class="ptCtl" data-for="holdings"><input class="drsc" id="ptSearch" placeholder="Search symbol or name..." oninput="filterPT()"/><select class="drsc catSel" id="ptCatFilter" onchange="filterPT()" title="Filter by timeframe"><option value="">All</option><option value="intraday">Intraday</option><option value="shortterm">Short-term</option><option value="longterm">Long-term</option></select></span>
    <span class="ptCtl" data-for="positions" style="display:none"><input class="drsc" id="posSearch" placeholder="Search symbol or name..." oninput="filterPositions()"/><select class="drsc catSel" id="posCatFilter" onchange="filterPositions()" title="Filter by timeframe"><option value="">All</option><option value="intraday">Intraday</option><option value="shortterm">Short-term</option><option value="longterm">Long-term</option></select></span>
    <span class="ptCtl" data-for="tradelog" style="display:none"><input class="drsc" id="tlSearch" placeholder="Search symbol or name..." oninput="filterTradeLog()"/><select class="drsc catSel" id="tlCatFilter" onchange="filterTradeLog()" title="Filter by timeframe"><option value="">All</option><option value="intraday">Intraday</option><option value="shortterm">Short-term</option><option value="longterm">Long-term</option></select></span>
  </div>

  <div class="pan ptBody" id="ptSub-holdings">
    <div id="ptSummary" class="ptSum"></div>
    <div class="scrollTbl">
    <table><thead><tr><th>#</th><th>Symbol</th><th>Type</th><th>Qty</th><th>Price</th><th>Date</th><th>LTP</th><th>Invested</th><th>Current Val.</th><th>P&amp;L</th><th>XIRR</th><th>Action</th><th></th></tr></thead>
      <tbody id="ptTblBdy"><tr class="emr"><td colspan="13">Loading...</td></tr></tbody>
    </table>
    </div>
  </div>

  <div class="pan ptBody" id="ptSub-positions" style="display:none">
    <div id="posSummary" class="ptSum"></div>
    <div class="scrollTbl">
    <table><thead><tr><th>#</th><th>Symbol</th><th>Type</th><th>Qty</th><th>Price</th><th>Date</th><th>LTP</th><th>Invested</th><th>Current Val.</th><th>P&amp;L</th><th>Action</th><th></th></tr></thead>
      <tbody id="posTblBdy"><tr class="emr"><td colspan="12">Loading...</td></tr></tbody>
    </table>
    </div>
  </div>

  <div class="pan ptBody" id="ptSub-tradelog" style="display:none">
    <div id="tlSummary" class="ptSum"></div>
    <div class="scrollTbl">
    <table><thead><tr><th>#</th><th>Symbol</th><th>Status</th><th>Category</th><th>Net Qty</th><th>Invested</th><th>Current/Exit Val.</th><th>P&amp;L</th><th>Return%</th><th>XIRR</th></tr></thead>
      <tbody id="tlTblBdy"><tr class="emr"><td colspan="10">Loading...</td></tr></tbody>
    </table>
    </div>
  </div>
</div>
<footer>
  <span class="fnote">This is for informational purposes only and not a investment advice. Always consult a qualified financial advisor. &middot; &copy; 2026 Roy. All rights reserved.</span>
  <div style="display:flex;align-items:center;gap:13px">
    <span class="favg" id="favg">Avg Return: <span class="v pos">-</span></span>
    <span class="fupd">Updated: <span id="ftm">-</span></span>
  </div>
</footer>
</div>

<button class="fab" onclick="openFB()">&#128172;</button>
<div class="mover" id="fbM">
  <div class="modal">
    <div class="mhd"><span class="mttl">Feedback</span><button class="mx" onclick="closeFB()">&#10005;</button></div>
    <div class="mbdy">
      <label class="mlbl">Type</label>
      <div class="mtyps"><button class="tybtn sel" onclick="selT(this)">Bug</button><button class="tybtn" onclick="selT(this)">Feature</button><button class="tybtn" onclick="selT(this)">Query</button><button class="tybtn" onclick="selT(this)">Other</button></div>
      <label class="mlbl">Your Email (for reply)</label><input class="minp" id="fbEmail" type="email" placeholder="your@email.com" style="height:34px"/>
      <label class="mlbl">Subject</label><input class="minp" id="fbSub" type="text" placeholder="Brief subject..." style="height:34px"/>
      <label class="mlbl">Message</label><textarea class="minp" id="fbMsg" rows="4" placeholder="Describe..."></textarea>
      <div class="mft"><button class="bcx" onclick="closeFB()">Cancel</button><button class="bok" id="fbSendBtn" onclick="sendFB()">Send</button></div>
    </div>
  </div>
</div>
<div id="toast"></div>
<script>
/* ══ STATE ══ */
var _st=[],_iRows=[],_ltRows=[],_newsMap={};
var _stF=null,_stDR=[],_iF=null,_iDR=[],_ltF=null,_ltDR=[];
var _stSort={key:null,dir:1},_iSort={key:null,dir:1},_ltSort={key:null,dir:1};
var _gDate="",_asOfTime="";
var _timer=null,_stBusy=false,_ltBusy=false,_tab="home";
var _lpOpen=false,_lpPinned=false,_lpTab="intraday",_lpTimer=null,_lpRows=[],_lpSrch=false;
var _aiHistory=[],_aiTmr=null,_sbT=null,_aiT=null,_aiPending=false;
var RFRSH=45000,_CIRC=2*Math.PI*38;
var _PL={st:null,intraday:null,lt:null,lpID:null,lpST:null,lpLT:null};
/* Per-tab independent period+price settings */
var _tabCfg={
  home:     {mode:"DAYS",val:1,pm:"ltp"},
  intraday: {mode:"DAYS",val:1,pm:"ltp"},
  shortterm:{mode:"DAYS",val:1,pm:"ltp"},
  longterm: {mode:"MONTHS",val:3,pm:"ltp"}
};
var _mode="DAYS",_val=1,_pm="ltp";
var _isMobile=window.matchMedia("(max-width:1024px)").matches;
var _fbType="Bug";

/* ══ HELPERS ══ */
function symLink(sym,name){
  var n=(name&&name.trim()&&name.trim()!==sym)?name.trim():"";
  return '<a href="https://www.screener.in/company/'+encodeURIComponent(sym)+'/consolidated" target="_blank"'+(n?' title="'+eh(n)+'"':'')+' class="sym-lnk">'+sym+'</a>';
}
var _curCategory="intraday";
var _wlSymSet=new Set();
function newsIco(sym){if(!_newsMap[sym])return"";return '<span class="nwi" data-sym="'+eh(sym)+'" title="News">&#9432;</span>';}
function noteIco(notes){var n=(notes||"").trim();if(!n)return"";return '<span class="noi" data-note="'+eh(n)+'" title="Note">&#128221;</span>';}
function isCheckFlagged_(sym,name){var re=/\u2713\s*$/;return re.test(String(sym||"").trim())||re.test(String(name||"").trim());}
function quickToggleWatchlist(sym,el){
  if(_wlSymSet.has(sym)){
    var row=_wlRows&&_wlRows.find(function(r){return r.symbol===sym;});
    if(!row)return;
    google.script.run.withSuccessHandler(function(r){if(r&&r.success){_wlSymSet.delete(sym);document.querySelectorAll('.wli[data-sym="'+sym+'"]').forEach(function(e){e.classList.remove("on");e.title="Add to watchlist";});if(document.querySelector(".tbtn[data-tab='watchlist']")&&document.getElementById("cnt-watchlist").style.display!=="none")loadWL(true);}}).withFailureHandler(function(){}).removeFromWatchlistFromWeb(_authToken,row.rowId);
  } else {
    google.script.run.withSuccessHandler(function(r){
      if(r&&r.sessionExpired){forceReauth();return;}
      if(r&&r.success){_wlSymSet.add(sym);document.querySelectorAll('.wli[data-sym="'+sym+'"]').forEach(function(e){e.classList.add("on");e.title="Remove from watchlist";});if(document.getElementById("cnt-watchlist").style.display!=="none")loadWL(true);}
    }).withFailureHandler(function(){}).addToWatchlistFromWeb(_authToken,sym,"");
  }
}
/* Diff-patch table body renderer. Fixes the flicker seen on every add/
   delete/periodic refresh in watchlist, portfolio and the order book: those
   used to rebuild the whole tbody's innerHTML on every render, which wipes
   and repaints every row even when only one changed. Instead: each logical
   item owns one or more <tr>s (tracked via data-key); unchanged rows are
   left completely untouched (checked via a data-sig hash of their own
   html), changed rows get only their own innerHTML updated, reordered rows
   are moved (not recreated), and only rows for removed items are deleted.
   items: [{key:string, rows:[htmlStr, ...]}]  — most callers use one row
   per item; portfolio uses two (main row + expandable trade-history row). */
/* Generic KPI drill-down search engine. The search box accepts free-text
   (matched against symbol/name, existing behaviour) PLUS space-separated
   structured tokens like "status:exited", "days>10", "cycles>=2",
   "return<-5", "exit>2026-08-01" — every recognised token must match
   (AND logic); unrecognised field names fall back to plain-text matching
   so nothing the user types ever gets silently dropped. `accessors` is a
   {fieldName: function(row){...}} map supplied per tab (ST/Intraday/LT
   have different underlying fields — e.g. "days" means tradingDays for
   ST/LT but tradingMins for Intraday). */
function kpiRowMatches_(row,query,accessors){
  var tokens=(query||"").trim().split(/\s+/).filter(Boolean);
  if(!tokens.length)return true;
  var textParts=[];
  for(var i=0;i<tokens.length;i++){
    var tk=tokens[i];
    var m=tk.match(/^([a-zA-Z]+)(>=|<=|>|<|:|=)(.+)$/);
    var acc=m?accessors[m[1].toLowerCase()]:null;
    if(!acc){textParts.push(tk);continue;}
    if(!kpiFieldMatch_(acc(row),m[2],m[3]))return false;
  }
  if(textParts.length){
    var q=textParts.join(" ").toUpperCase();
    var hay=((row.symbol||"")+" "+(row.name||"")).toUpperCase();
    if(hay.indexOf(q)<0)return false;
  }
  return true;
}
function kpiFieldMatch_(rawVal,op,val){
  if(rawVal==null||rawVal==="")return false;
  if(op===":"||op==="="){
    return String(rawVal).toUpperCase().indexOf(String(val).toUpperCase())>=0;
  }
  var num=Number(rawVal),target=Number(val);
  if(!isNaN(num)&&!isNaN(target)){
    if(op===">")return num>target; if(op==="<")return num<target;
    if(op===">=")return num>=target; if(op==="<=")return num<=target;
  }
  var dt=new Date(rawVal).getTime(),dv=new Date(val).getTime();
  if(!isNaN(dt)&&!isNaN(dv)){
    if(op===">")return dt>dv; if(op==="<")return dt<dv;
    if(op===">=")return dt>=dv; if(op==="<=")return dt<=dv;
  }
  return false;
}
/* Generic clickable-column-header sort. state:{key:null,dir:1}; accessors
   is the same shape as above. Ascending on first click of a column,
   flips on repeat clicks. Non-comparable (null) values always sort last
   regardless of direction. */
function kpiSortRows_(rows,state,accessors){
  if(!state.key||!accessors[state.key])return rows;
  var acc=accessors[state.key],dir=state.dir;
  return rows.slice().sort(function(a,b){
    var av=acc(a),bv=acc(b);
    var an=av==null||av==="", bn=bv==null||bv==="";
    if(an&&bn)return 0; if(an)return 1; if(bn)return -1;
    if(typeof av==="string"||typeof bv==="string"){var as=String(av),bs=String(bv);return as<bs?-dir:as>bs?dir:0;}
    return (av-bv)*dir;
  });
}
function kpiSortClick_(theadId,key,state,rerender){
  state.dir=(state.key===key)?-state.dir:1;
  state.key=key;
  document.querySelectorAll("#"+theadId+" th[data-col]").forEach(function(th){
    th.classList.remove("srt-asc","srt-desc");
    if(th.getAttribute("data-col")===key)th.classList.add(state.dir>0?"srt-asc":"srt-desc");
  });
  rerender();
}
/* Parses a "<tr ...>...</tr>" string into a real <tr> DOM node. This is
   required because assigning such a string to a <tr> element's innerHTML
   is invalid HTML nesting — a <tr> cannot contain another <tr>, so
   browsers silently strip the *outer* <tr> tag (and, critically, any
   id/style/attributes on it) while still inserting its <td> children.
   For plain rows that went unnoticed (nothing read the wrapper's own
   attributes). For the portfolio expand/collapse row, whose own <tr>
   carries id="ptExp_SYMBOL" and style="display:none", it meant that id
   and that style were silently discarded on every render — so the row
   could never truly be found or hidden, regardless of _ptExpanded state. */
function parseTrHtml_(html){
  var tmp=document.createElement("table");
  tmp.innerHTML="<tbody>"+html+"</tbody>";
  return tmp.querySelector("tr");
}
function patchTbody(tbody,items){
  if(!tbody)return;
  var groups={};
  Array.prototype.forEach.call(tbody.children,function(tr){
    var k=tr.getAttribute("data-key")||"";
    (groups[k]=groups[k]||[]).push(tr);
  });
  var usedKeys={},ref=null;
  (items||[]).forEach(function(item){
    usedKeys[item.key]=true;
    var nodes=groups[item.key]||[];
    (item.rows||[]).forEach(function(html,idx){
      var tr=nodes[idx];
      if(tr){
        if(tr.getAttribute("data-sig")!==html){
          var fresh=parseTrHtml_(html);
          fresh.setAttribute("data-key",item.key);
          fresh.setAttribute("data-sig",html);
          tr.parentNode.replaceChild(fresh,tr);
          nodes[idx]=fresh; tr=fresh;
        }
      } else {
        tr=parseTrHtml_(html);
        tr.setAttribute("data-key",item.key);
        tr.setAttribute("data-sig",html);
      }
      var wantedPos=ref?ref.nextSibling:tbody.firstChild;
      if(wantedPos!==tr)tbody.insertBefore(tr,wantedPos);
      ref=tr;
    });
    for(var j=(item.rows||[]).length;j<nodes.length;j++){if(nodes[j]&&nodes[j].parentNode)nodes[j].parentNode.removeChild(nodes[j]);}
  });
  Object.keys(groups).forEach(function(k){
    if(!usedKeys[k])groups[k].forEach(function(n){if(n.parentNode)n.parentNode.removeChild(n);});
  });
}
function categoryBdg_(category){
  if(category==="intraday")return '<span class="cbdg" style="background:rgba(60,200,220,.15);color:var(--cy)">Intraday</span>';
  if(category==="shortterm")return '<span class="cbdg" style="background:rgba(70,200,140,.15);color:var(--gr)">SHORT</span>';
  return '<span class="cbdg" style="background:rgba(255,180,60,.15);color:var(--am)">LONG</span>';
}
function symCell(sym,name,cons,catOverride,notes,flagCheck){
  var cat=catOverride!==undefined?catOverride:_curCategory;
  var inWl=_wlSymSet.has(sym);
  var wlIco='<span class="wli'+(inWl?" on":"")+'" data-sym="'+eh(sym)+'" onclick="event.stopPropagation();quickToggleWatchlist(\''+eh(sym)+'\',this)" title="'+(inWl?"Remove from watchlist":"Add to watchlist")+'">&#9733;</span>';
  var nwIco=newsIco(sym);
  var noIco=noteIco(notes);
  var flCls=(flagCheck&&isCheckFlagged_(sym,name))?" symfl":"";
  // Order-book icon removed from symCell — accessible only from the live panel
  return '<td class="sym'+flCls+'">'+wlIco+symLink(sym,name||"")+((nwIco||noIco)?('<span class="symacts">'+nwIco+noIco+'</span>'):'')+(cons?'<span class="cbdg">x'+cons+'</span>':'')+'</td>';
}
function fN(n){if(n===null||n===undefined||n===""||n!==n)return"-";var v=Number(n);return isNaN(v)||v===0?"-":v.toLocaleString("en-IN",{maximumFractionDigits:2});}
function fCr(n){if(n===null||n===undefined||n===""||n!==n)return"-";var v=Number(n);if(isNaN(v)||v===0)return"-";return(v/1e7).toLocaleString("en-IN",{maximumFractionDigits:2});}
function fNZ(n){if(n===null||n===undefined||n===""||n!==n)return"-";var v=Number(n);return isNaN(v)?"-":v.toLocaleString("en-IN",{maximumFractionDigits:2});}
function fR(n){if(n===null||n===undefined||n!==n)return"-";var v=Number(n);return isNaN(v)?"-":(v>=0?"+":"")+(v*100).toFixed(2)+"%";}
function hC(v){var n=Number(v);return(v!==null&&v!==undefined&&v!==""&&!isNaN(n)&&n>0)?'<span style="color:var(--cy);font-size:11px">'+fNZ(v)+'</span>':'-';}
function eh(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}
function fRS(s){if(!s||!s.trim())return"-";return s.replace(/\d+(\.\d+)?/g,function(m){var n=parseFloat(m);return isNaN(n)?m:n.toFixed(2);});}
function oPill(o){var s=String(o||"").toUpperCase();if(s.startsWith("B"))return'<span class="pill pbuy">BUY</span>';if(s.startsWith("S"))return'<span class="pill psell">SELL</span>';return'<span class="pill prun">'+o+'</span>';}
function stPill(r){if(r.exitFlag==="EXITED")return stPr(r)?'<span class="pill psuc">SUCCESS</span>':'<span class="pill pfal">FAILED</span>';return'<span class="pill prun">RUNNING</span>';}
function iPill(r){if(isAfterMkt()&&r.exitFlag!=="EXITED")return'<span class="pill pcls">CLOSED</span>';if(r.exitFlag==="EXITED")return iPr(r)?'<span class="pill psuc">SUCCESS</span>':'<span class="pill pfal">FAILED</span>';return'<span class="pill prun">RUNNING</span>';}
function ltPill(r){if(r.exitFlag==="EXITED")return ltR(r)>0?'<span class="pill psuc">SUCCESS</span>':'<span class="pill pfal">FAILED</span>';return'<span class="pill prun">RUNNING</span>';}
function upd(id,html){var e=document.getElementById(id);if(e)e.innerHTML=html;}/* no opacity change */
function toast(msg,ms){ms=ms||2600;var e=document.getElementById("toast");e.textContent=msg;e.classList.add("show");setTimeout(function(){e.classList.remove("show");},ms);}
function setK(i,v,total){
  var e=document.getElementById("k"+i);if(e)e.textContent=(v===null||v===undefined||v==="")?"–":v;
  var rc=document.getElementById("k"+i+"rc"),rp=document.getElementById("k"+i+"rp");
  var ratio=(total&&Number(total)>0&&v!=null&&v!=="")?Math.max(0,Math.min(1,Number(v)/Number(total))):0;
  if(rc){
    var circ=2*Math.PI*15;
    rc.style.strokeDasharray=circ.toFixed(2);
    rc.style.strokeDashoffset=(circ*(1-ratio)).toFixed(2);
  }
  if(rp)rp.textContent=Math.round(ratio*100)+"%";
}
function clrKPIs(){for(var i=0;i<6;i++)setK(i,"-",0);}
function clearFooter(){var sp=document.querySelector("#favg .v");if(sp){sp.textContent="-";sp.className="v";}document.getElementById("ftm").textContent="-";}
function updFoot(rows,retFn){var wp=rows.filter(function(r){return r.ltp!=null;});var avg=wp.length?wp.reduce(function(s,r){return s+(retFn(r)||0);},0)/wp.length:null;var sp=document.querySelector("#favg .v");if(sp){sp.textContent=avg!=null?fR(avg):"-";sp.className="v "+(avg!=null&&avg>=0?"pos":"neg");}document.getElementById("ftm").textContent=new Date().toLocaleTimeString("en-IN",{hour12:false});}
function animD(arcId,pctId,rate,clr){var arc=document.getElementById(arcId),pe=document.getElementById(pctId);if(!arc||!pe)return;var h=clr||(rate>.5?"#00a362":rate>.3?"#d4800a":"#e02040");arc.setAttribute("stroke-dasharray",(rate*_CIRC)+" "+((1-rate)*_CIRC));arc.style.stroke=h;pe.textContent=Math.round(rate*100)+"%";pe.style.color=h;}
function renderBars(secId,items){var el=document.getElementById(secId);if(!el)return;var tot=Math.max(items.reduce(function(s,it){return s+it.v;},0),1);el.innerHTML=items.map(function(it){var w=it.v/tot*100;return'<div class="brow"><span class="blbl">'+it.l+'</span><div class="btrk"><div class="bfil" style="background:'+it.c+'" data-w="'+w+'"></div></div><span class="bnum" style="color:'+it.c+'">'+it.v+'</span></div>';}).join("");requestAnimationFrame(function(){el.querySelectorAll(".bfil").forEach(function(e){e.style.width=e.dataset.w+"%";});});}

/* ══ DISCLAIMER ══ */


/* ══ SIDEBAR ══ */
if(!_isMobile){
  document.getElementById("sbT").style.display="block";
  document.getElementById("sbT").addEventListener("mouseenter",openSB);
  document.getElementById("sb").addEventListener("mouseenter",function(){clearTimeout(_sbT);});
  document.getElementById("sb").addEventListener("mouseleave",function(){_sbT=setTimeout(closeSB,380);});
}
function toggleSB(){document.getElementById("sb").classList.contains("open")?closeSB():openSB();}
function openSB(){clearTimeout(_sbT);document.getElementById("sb").classList.add("open");document.getElementById("sbO").classList.add("show");}
function closeSB(){document.getElementById("sb").classList.remove("open");document.getElementById("sbO").classList.remove("show");}
document.getElementById("sbO").addEventListener("click",closeSB);
document.getElementById("sbO").addEventListener("touchstart",function(e){e.preventDefault();closeSB();},{passive:false});

/* ══ AI PANEL — 1s hover delay on desktop, stays open (no auto-hide) ══ */
if(!_isMobile){
  var _aiTEl=document.getElementById("aiT");
  _aiTEl.style.display="block";
  _aiTEl.addEventListener("mouseenter",function(){_aiT=setTimeout(openAI_p,1000);});
  /* Previously had no cancel-on-leave: a brief graze across this edge strip
     (e.g. moving toward the native scrollbar, which sits in the same pixels)
     armed a 1s timer that fired even after the cursor had long since moved
     away, popping the panel open mid-scroll. Both of these cancel it. */
  _aiTEl.addEventListener("mouseleave",function(){clearTimeout(_aiT);});
  window.addEventListener("wheel",function(){clearTimeout(_aiT);},{passive:true});
  document.getElementById("aip").addEventListener("mouseenter",function(){clearTimeout(_aiT);});
  /* No mouseleave handler — panel stays open until user explicitly closes */
}
var _aiPinned=false;
function toggleAIPin(){_aiPinned=!_aiPinned;var btn=document.getElementById("aiPinBtn");if(btn){btn.style.color=_aiPinned?"var(--ac)":"";btn.style.borderColor=_aiPinned?"var(--ac)":"";}}
function toggleAI(){document.getElementById("aip").classList.contains("open")?closeAI():openAI_p();}
function openAI_p(){clearTimeout(_aiT);document.getElementById("aip").classList.add("open");document.getElementById("aiO").classList.add("show");}
function closeAI(){if(_aiPinned)return;document.getElementById("aip").classList.remove("open");document.getElementById("aiO").classList.remove("show");}
document.getElementById("aiO").addEventListener("click",function(){if(!_aiPinned)closeAI();});
document.getElementById("aiO").addEventListener("touchstart",function(e){e.preventDefault();if(!_aiPinned)closeAI();},{passive:false});
(function(){
  var panel=document.getElementById("aip"),handle=document.getElementById("aipResize");
  if(!panel||!handle)return;
  var dragging=false,startX=0,startWidth=0,MIN_W=320;
  var saved=Number(localStorage.getItem("aipWidth"))||0;
  if(saved)panel.style.width=Math.min(Math.max(saved,MIN_W),Math.min(900,window.innerWidth-60))+"px";
  function pos(e){return e.touches?e.touches[0].clientX:e.clientX;}
  function onDown(e){dragging=true;startX=pos(e);startWidth=panel.getBoundingClientRect().width;handle.classList.add("dragging");document.body.style.userSelect="none";e.preventDefault();}
  function onMove(e){
    if(!dragging)return;
    var maxW=Math.min(900,window.innerWidth-60);
    var newW=Math.min(Math.max(startWidth+(startX-pos(e)),MIN_W),maxW);
    panel.style.width=newW+"px";
  }
  function onUp(){
    if(!dragging)return;
    dragging=false;handle.classList.remove("dragging");document.body.style.userSelect="";
    try{localStorage.setItem("aipWidth",parseInt(panel.style.width,10));}catch(e){}
  }
  handle.addEventListener("mousedown",onDown);
  handle.addEventListener("touchstart",onDown,{passive:false});
  document.addEventListener("mousemove",onMove);
  document.addEventListener("touchmove",onMove,{passive:false});
  document.addEventListener("mouseup",onUp);
  document.addEventListener("touchend",onUp);
})();

/* ══ AI CHAT ══ */
/* Lightweight markdown renderer for AI chat responses. Escapes the raw
   text FIRST (so any stray HTML-like content in the model's output is
   neutralized), then converts markdown syntax operating only on the
   already-escaped ASCII punctuation (#, *, |, `, [ ]) into safe HTML.
   Covers what a financial-analysis chat actually produces: tables,
   headings, bold/italic, inline code, links, and lists. */
function mdInline_(s){
  s = s.replace(/\*\*(.+?)\*\*/g, "<b>$1</b>");
  s = s.replace(/\*([^*\n]+?)\*/g, "<i>$1</i>");
  s = s.replace(/`([^`]+?)`/g, "<code>$1</code>");
  s = s.replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, '<a href="$2" target="_blank" rel="noopener" style="color:var(--ac)">$1</a>');
  return s;
}
function mdToHtml_(text){
  var raw = eh(String(text||""));
  var lines = raw.split(/\n/);
  var out = [], i = 0;
  function isTableRow(l){ return /^\s*\|.*\|\s*$/.test(l); }
  function isTableSep(l){ return /^\s*\|?[\s:-]*-{2,}[\s:|-]*\|?\s*$/.test(l); }
  function splitRow(l){
    return l.trim().replace(/^\|/,"").replace(/\|$/,"").split("|").map(function(c){return mdInline_(c.trim());});
  }
  while(i<lines.length){
    var line = lines[i];
    if (isTableRow(line) && i+1<lines.length && isTableSep(lines[i+1])){
      var header = splitRow(line); i += 2;
      var rows = [];
      while(i<lines.length && isTableRow(lines[i])){ rows.push(splitRow(lines[i])); i++; }
      var thead = "<thead><tr>"+header.map(function(h){return "<th>"+h+"</th>";}).join("")+"</tr></thead>";
      var tbody = "<tbody>"+rows.map(function(r){return "<tr>"+r.map(function(c){return "<td>"+c+"</td>";}).join("")+"</tr>";}).join("")+"</tbody>";
      out.push('<div class="ai-tblwrap"><table class="ai-tbl">'+thead+tbody+'</table></div>');
      continue;
    }
    var hm = line.match(/^(#{1,6})\s+(.*)$/);
    if (hm){ out.push('<div class="ai-h ai-h'+hm[1].length+'">'+mdInline_(hm[2])+'</div>'); i++; continue; }
    if (/^\s*[-*]\s+/.test(line)){
      var uitems=[];
      while(i<lines.length && /^\s*[-*]\s+/.test(lines[i])){ uitems.push(mdInline_(lines[i].replace(/^\s*[-*]\s+/,""))); i++; }
      out.push("<ul>"+uitems.map(function(it){return "<li>"+it+"</li>";}).join("")+"</ul>");
      continue;
    }
    if (/^\s*\d+\.\s+/.test(line)){
      var oitems=[];
      while(i<lines.length && /^\s*\d+\.\s+/.test(lines[i])){ oitems.push(mdInline_(lines[i].replace(/^\s*\d+\.\s+/,""))); i++; }
      out.push("<ol>"+oitems.map(function(it){return "<li>"+it+"</li>";}).join("")+"</ol>");
      continue;
    }
    if (/^\s*$/.test(line)){ i++; continue; }
    out.push('<div class="ai-p">'+mdInline_(line)+'</div>');
    i++;
  }
  return out.join("");
}
function addAIMsg(role,text){
  var bd=document.getElementById("aiBd");
  var w=bd.querySelector(".ai-welcome");if(w)w.remove();
  var d=document.createElement("div");d.className="ai-msg "+role;
  var b=document.createElement("div");b.className="ai-bubble";
  b.innerHTML = (role==="bot") ? mdToHtml_(text) : eh(text).replace(/\n/g,"<br/>");
  d.appendChild(b);bd.appendChild(d);bd.scrollTop=bd.scrollHeight;
}
function showDots(){var bd=document.getElementById("aiBd");if(document.getElementById("aiDts"))return;var d=document.createElement("div");d.className="ai-msg bot";d.id="aiDts";d.innerHTML='<div class="ai-dots"><span></span><span></span><span></span></div>';bd.appendChild(d);bd.scrollTop=bd.scrollHeight;}
function removeDots(){var e=document.getElementById("aiDts");if(e)e.remove();}
function askAI(q){openAI_p();document.getElementById("aiIn").value=q;sendAI();}
/* AI gets its OWN broad, unfiltered ST/LT data — independent of whatever
   period filter you currently have selected on screen. _st/_ltRows get
   overwritten by the active display filter, so the AI was only ever seeing
   that narrow slice. Intraday stays scoped to today's session by design
   (per original spec — same-day data only). Fetched lazily, only when AI
   is actually used, cached 90s — no new background polling added. */
var _aiBroadST=null,_aiBroadLT=null,_aiBroadTS=0;
function ensureBroadAI(cb){
  if(_aiBroadST&&_aiBroadLT&&(Date.now()-_aiBroadTS)<90000){cb();return;}
  var pending=2;
  function done(){if(--pending===0){_aiBroadTS=Date.now();cb();}}
  google.script.run.withSuccessHandler(function(r){_aiBroadST=r.rows||[];done();}).withFailureHandler(function(){_aiBroadST=_aiBroadST||_st||[];done();}).runAnalysisFromWeb();
  google.script.run.withSuccessHandler(function(r){_aiBroadLT=r.rows||[];done();}).withFailureHandler(function(){_aiBroadLT=_aiBroadLT||_ltRows||[];done();}).getLTAnalysisDataFromWeb();
}
function stripHtml_(s){return String(s).replace(/<[^>]*>/g," ").replace(/&amp;/g,"&").replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/&quot;/g,'"').replace(/\s+/g," ").trim();}
function newsSection_(){
  var keys=Object.keys(_newsMap||{});
  if(!keys.length)return "";
  var order=[];
  [_aiBroadST,_iRows,_aiBroadLT].forEach(function(arr){(arr||[]).forEach(function(r){if(_newsMap[r.symbol]&&order.indexOf(r.symbol)<0)order.push(r.symbol);});});
  keys.forEach(function(k){if(order.indexOf(k)<0)order.push(k);});
  var lines=order.slice(0,15).map(function(sym){return sym+": "+stripHtml_(_newsMap[sym]).slice(0,140);});
  return "=== NEWS (from symbol hover-icon source, "+keys.length+" symbols total, showing "+lines.length+") ===\n"+lines.join("\n")+"\n\n";
}
function buildSysCtx(){
  var ctx="You are a specialized AI assistant for a personal NSE/BSE India stock signal dashboard with three categories: Intraday (same-day session), Short-term (multi-day cycles), Long-term (holdings).\n";
  ctx+="RULES: Only suggest/discuss symbols that appear in the data below — never invent stocks outside it. If a symbol's price isn't shown (compact rows below only give return%), say you don't have its current price rather than estimating or guessing one — the person can check the live panel for it. Never ask what stocks are in the tool — you already have them below. Use multi-day data for ST/LT and today's session for Intraday. Answer in INR (Rs), be specific and decisive.\n\n";
  var now=new Date();
  ctx+="Date/Time (IST): "+now.toLocaleString("en-IN")+"\n\n";

  /* Groq's free-tier TPM limit (12000 tokens/minute) was being exceeded
     because every row was sent in full verbose detail regardless of how
     many signals exist. This stays adaptive: small lists keep full detail
     (entry/SL/target etc.), large lists fall back to a compact ranked
     symbol:return%:days line capped at a fixed count — so total size stays
     bounded no matter how many signals the dashboard is tracking. */
  var FULL_CAP=18, COMPACT_TOP=14, COMPACT_BOTTOM=8;

  function section(rows,retFn,label,fullLine){
    if(!rows||!rows.length)return "=== "+label+" === (none loaded)\n\n";
    var withRet=rows.map(function(r){return{r:r,ret:retFn(r)};});
    withRet.sort(function(a,b){return(b.ret==null?-1e9:b.ret)-(a.ret==null?-1e9:a.ret);});
    var out="=== "+label+" ("+rows.length+" total) ===\n";
    if(rows.length<=FULL_CAP){
      withRet.forEach(function(x){out+=fullLine(x.r,x.ret)+"\n";});
    } else {
      out+="Showing top "+COMPACT_TOP+" + bottom "+COMPACT_BOTTOM+" by return% (most relevant for ranking), with live price:\n";
      var picked=withRet.slice(0,COMPACT_TOP).concat(withRet.slice(-COMPACT_BOTTOM));
      var seen={};
      picked.forEach(function(x){
        if(seen[x.r.symbol])return;seen[x.r.symbol]=1;
        var ret=x.ret,retStr=ret!=null?(ret>=0?"+":"")+(ret*100).toFixed(1)+"%":"-";
        var ltpStr=x.r.ltp!=null?fNZ(x.r.ltp):"-";
        out+=x.r.symbol+":Rs"+ltpStr+":"+retStr+(x.r.tradingDays?","+x.r.tradingDays+"d":"")+(x.r.exitFlag==="EXITED"?",EXITED":",RUNNING")+" | ";
      });
      out=out.replace(/ \| $/,"\n");
    }
    return out+"\n";
  }

  ctx+=section(_aiBroadST||_st,stR,"SHORT-TERM SIGNALS (full history, not limited to your current screen filter)",function(r,ret){
    var retStr=ret!=null?(ret>=0?"+":"")+(ret*100).toFixed(2)+"%":"-";
    var slV=r.sl||r.sheetSL||0,tgV=r.target||r.sheetTarget||0;
    return r.symbol+" | Entry:"+fNZ(r.entry)+" | LTP:"+(r.ltp?fNZ(r.ltp):"-")+" | Ret:"+retStr+" | SL:"+(slV?fNZ(slV):"-")+" | Tgt:"+(tgV?fNZ(tgV):"-")+" | "+(r.tradingDays||0)+"d | "+(r.exitFlag==="EXITED"?"EXITED":"RUNNING");
  });
  ctx+=section(_iRows,iR,"INTRADAY SIGNALS (today's session only — by design)",function(r,ret){
    var retStr=ret!=null?(ret>=0?"+":"")+(ret*100).toFixed(2)+"%":"-";
    return r.symbol+" | "+(r.isSell?"SELL":"BUY")+" | Entry:"+fNZ(r.entry)+" | LTP:"+(r.ltp?fNZ(r.ltp):"-")+" | Ret:"+retStr+" | Tgt:"+(r.target?fNZ(r.target):"-")+" | "+(r.tradingMins||0)+"min | "+(r.exitFlag==="EXITED"?"EXITED":"RUNNING");
  });
  ctx+=section(_aiBroadLT||_ltRows,ltR,"LONG-TERM HOLDINGS (full history, not limited to your current screen filter)",function(r,ret){
    var retStr=ret!=null?(ret>=0?"+":"")+(ret*100).toFixed(2)+"%":"-";
    return r.symbol+" | Entry:"+fNZ(r.entry)+" | LTP:"+(r.ltp?fNZ(r.ltp):"-")+" | Ret:"+retStr+" | "+(r.tradingDays||0)+"d | "+(r.exitFlag==="EXITED"?"EXITED":"RUNNING");
  });
  ctx+=newsSection_();

  ctx+="When asked to suggest stocks: rank by return%, momentum, and risk. Only name symbols from the data above.\n";
  return ctx;
}
function sendAI(){
  if(_aiPending)return;
  var inp=document.getElementById("aiIn"),btn=document.getElementById("aiSend");
  var msg=inp.value.trim();if(!msg)return;
  inp.value="";inp.style.height="auto";btn.disabled=true;_aiPending=true;
  addAIMsg("user",msg);showDots();
  _aiHistory.push({role:"user",content:msg});
  if(_aiHistory.length>16)_aiHistory=_aiHistory.slice(-16);
  /* Timeout: show error (not local fallback) after 35s */
  clearTimeout(_aiTmr);
  _aiTmr=setTimeout(function(){
    removeDots();btn.disabled=false;_aiPending=false;
    addAIMsg("bot","Request timed out. The AI may be busy — please try again.");
  },35000);
  ensureBroadAI(function(){
    google.script.run
      .withSuccessHandler(function(r){
        clearTimeout(_aiTmr);removeDots();btn.disabled=false;_aiPending=false;
        if(r&&r.text){
          _aiHistory.push({role:"assistant",content:r.text});
          addAIMsg("bot",r.text);
        } else if(r&&r.quota){
          /* Groq daily rate limit hit */
          addAIMsg("bot","Groq rate limit reached for today. The AI will be available again tomorrow. Your Groq API resets daily.\n\nError: "+(r.error||""));
        } else {
          addAIMsg("bot","Error: "+(r&&r.error?r.error:"No response from server."));
        }
      })
      .withFailureHandler(function(e){
        clearTimeout(_aiTmr);removeDots();btn.disabled=false;_aiPending=false;
        addAIMsg("bot","Server error: "+(e&&e.message?e.message:"Check Apps Script logs."));
      })
      .processGroqAIFromWeb(_aiHistory.slice(-4),buildSysCtx());
  });
}

/* ══ LIVE PANEL ══ */
function openLP(type){_lpOpen=true;document.getElementById("lp").classList.add("open");document.getElementById("lpO").classList.add("show");document.querySelectorAll(".sbb[data-lp]").forEach(function(b){b.classList.toggle("lon",b.dataset.lp===type);});switchLP(type);startLPRefresh();}
function closeLP(){if(_lpPinned)return;_lpOpen=false;document.getElementById("lp").classList.remove("open");document.getElementById("lpO").classList.remove("show");document.querySelectorAll(".sbb[data-lp]").forEach(function(b){b.classList.remove("lon");});}
function toggleLPPin(){_lpPinned=!_lpPinned;var btn=document.getElementById("lpPinBtn");if(btn){btn.style.color=_lpPinned?"var(--ac)":"";btn.style.borderColor=_lpPinned?"var(--ac)":"";}}
document.getElementById("lpO").addEventListener("click",function(){if(!_lpPinned)closeLP();});
document.getElementById("lpO").addEventListener("touchstart",function(e){e.preventDefault();if(!_lpPinned)closeLP();},{passive:false});
function switchLP(type){
  _lpTab=type;
  var cls={intraday:"la-id",shortterm:"la-st",longterm:"la-lt"};
  document.querySelectorAll(".lptab").forEach(function(b){b.className="lptab";if(b.dataset.lt===type)b.classList.add(cls[type]||"");});
  document.getElementById("lpTtl").textContent={intraday:"Intraday Live",shortterm:"Short-term Live",longterm:"Long-term Live"}[type]||"Live";
  var ck={intraday:"lpID",shortterm:"lpST",longterm:"lpLT"}[type];
  if(ck&&_PL[ck]&&_PL[ck].length){_lpRows=_PL[ck];renderLP(_lpRows,type);}
  else{document.getElementById("lpBdy").innerHTML='<div style="text-align:center;padding:30px;color:var(--dm);font-family:\'Barlow Condensed\',sans-serif;letter-spacing:2px;font-size:10px">LOADING...</div>';fetchLPBg(type);}
}
function startLPRefresh(){clearInterval(_lpTimer);_lpTimer=setInterval(function(){fetchLPBg(_lpTab);},15000);}
function fetchLPBg(type){
  google.script.run
    .withSuccessHandler(function(r){var rows=r.rows||[];var ck={intraday:"lpID",shortterm:"lpST",longterm:"lpLT"}[type];if(ck)_PL[ck]=rows;if(_lpTab===type){_lpRows=rows;var q=document.getElementById("lpSIn").value;renderLP(q?_lpRows.filter(function(x){var ql=q.toLowerCase();return(x.symbol||"").toLowerCase().includes(ql)||(x.name||"").toLowerCase().includes(ql);}):rows,type);}document.getElementById("lpTm").textContent=new Date().toLocaleTimeString("en-IN",{hour12:false});})
    .withFailureHandler(function(){})
    .getLiveSignalsFromWeb(type);
}
function toggleLPSearch(){_lpSrch=!_lpSrch;document.getElementById("lpSI").classList.toggle("on",_lpSrch);document.getElementById("lpSW").classList.toggle("on",_lpSrch);if(_lpSrch)setTimeout(function(){document.getElementById("lpSIn").focus();},300);else{document.getElementById("lpSIn").value="";renderLP(_lpRows,_lpTab);}}
function filterLP(){var q=document.getElementById("lpSIn").value.toLowerCase();renderLP(q?_lpRows.filter(function(r){return(r.symbol||"").toLowerCase().includes(q)||(r.name||"").toLowerCase().includes(q);}):_lpRows,_lpTab);}
function renderLP(rows,type){
  var bdy=document.getElementById("lpBdy");
  if(!rows.length){bdy.innerHTML='<div style="text-align:center;padding:30px;color:var(--dm);font-family:\'Barlow Condensed\',sans-serif;letter-spacing:2px;font-size:10px">NO SIGNALS</div>';return;}
  var h="",b="";
  if(type==="shortterm"){
    h="<tr><th></th><th>Symbol</th><th>LTP</th><th>SL</th><th>Target</th></tr>";
    b=rows.map(function(r){var ob='<td><span class="obi" onclick="openOrderModal(\''+r.symbol+'\',\'shortterm\','+(r.ltp||0)+')" title="Place order" style="cursor:pointer;font-size:12px;color:var(--ac)">&#9654;</span></td>';return'<tr>'+ob+symCell(r.symbol,r.name,"","shortterm",null,true)+'<td>'+fNZ(r.ltp)+'</td><td style="color:var(--rd)">'+fNZ(r.sl)+'</td><td style="color:var(--gr)">'+fNZ(r.target)+'</td></tr>';}).join("");
  } else if(type==="longterm"){
    h="<tr><th></th><th>Symbol</th><th>LTP</th><th>SL</th><th>Target</th></tr>";
    b=rows.map(function(r){var ob='<td><span class="obi" onclick="openOrderModal(\''+r.symbol+'\',\'longterm\','+(r.ltp||0)+')" title="Place order" style="cursor:pointer;font-size:12px;color:var(--ac)">&#9654;</span></td>';return'<tr>'+ob+symCell(r.symbol,r.name,"","longterm",null,true)+'<td>'+fNZ(r.ltp)+'</td><td style="color:var(--rd);font-weight:600">'+(r.sl?fNZ(r.sl):'-')+'</td><td style="color:var(--gr)">'+(r.target?fNZ(r.target):'-')+'</td></tr>';}).join("");
  } else {
    h="<tr><th></th><th>Symbol</th><th>Order</th><th>LTP</th><th>Open</th><th>High</th><th>SL</th><th>Target</th><th>Resi/Supp</th></tr>";
    b=rows.map(function(r){var ob='<td><span class="obi" onclick="openOrderModal(\''+r.symbol+'\',\'intraday\','+(r.ltp||r.entry||0)+')" title="Place order" style="cursor:pointer;font-size:12px;color:var(--ac)">&#9654;</span></td>';return'<tr>'+ob+symCell(r.symbol,r.name,"","intraday",null,true)+'<td>'+oPill(r.order)+'</td><td>'+fNZ(r.entry)+'</td><td style="color:var(--dm)">'+fNZ(r.open)+'</td><td>'+fNZ(r.high)+'</td><td style="color:var(--rd);font-weight:600">'+(r.sl?fNZ(r.sl):'-')+'</td><td style="color:var(--gr)">'+fNZ(r.target)+'</td><td style="color:var(--dm);font-size:10px">'+fRS(r.resiSupp)+'</td></tr>';}).join("");
  }
  bdy.innerHTML='<table><thead>'+h+'</thead><tbody>'+b+'</tbody></table>';
}

/* ══ TICKERS ══ */
function fetchTicker(){google.script.run.withSuccessHandler(function(ix){if(!ix||!ix.length)return;[{p:"tSX",c:"tSXc",d:ix[0]},{p:"tNF",c:"tNFc",d:ix[1]},{p:"tBN",c:"tBNc",d:ix[2]}].forEach(function(t){if(!t.d)return;var pe=document.getElementById(t.p),ce=document.getElementById(t.c);if(pe)pe.textContent=fNZ(t.d.price);if(ce){var ch=t.d.chg;ce.textContent=(ch>=0?"+":"")+ch.toFixed(2)+"%";ce.className="htck-c "+(ch>=0?"pos":"neg");}});}).withFailureHandler(function(){}).getMarketIndicesFromWeb();}
setInterval(fetchTicker,300000);
function fetchNews(){google.script.run.withSuccessHandler(function(m){_newsMap=m||{};if(_PL.st&&_tab==="shortterm")renderST(_PL.st);if(_PL.intraday&&_tab==="intraday")renderIntraday(_PL.intraday);if(_PL.lt&&_tab==="longterm")renderLT(_PL.lt);}).withFailureHandler(function(){}).getNewsDataFromWeb();}

/* News tooltip — floating element appended to body, positioned via JS so it
   always escapes ancestor overflow:hidden (table panels, drawers) and never
   gets clipped. Scrolls internally if content is long. Pinned by click for
   touch devices since they have no real hover.
   Hide uses a short grace-period timer (not instant) because the mouse has
   to travel from the icon to the box — hiding immediately on mouseout (the
   old behaviour) closed the box before the cursor ever reached it, so links
   inside could never be clicked. */
(function(){
  var box=document.getElementById("nwBox"),pinned=false,hideT=null;
  function place(ic){
    var r=ic.getBoundingClientRect(),bw=box.offsetWidth||280;
    box.style.left=Math.max(8,Math.min(r.left,window.innerWidth-bw-8))+"px";
    var top=r.top-box.offsetHeight-8;
    box.style.top=(top<8?r.bottom+8:top)+"px";
  }
  function showFor(ic){
    var sym=ic.getAttribute("data-sym"),html=_newsMap[sym];if(!html)return;
    clearTimeout(hideT);
    box.innerHTML=html;box.classList.add("show");place(ic);
  }
  function scheduleHide(){
    if(pinned)return;
    clearTimeout(hideT);
    hideT=setTimeout(function(){if(!pinned)box.classList.remove("show");},300);
  }
  document.addEventListener("mouseover",function(e){
    var ic=e.target.closest&&e.target.closest(".nwi");
    if(ic){showFor(ic);return;}
    if(box.contains(e.target))clearTimeout(hideT); // cursor reached the box itself — cancel pending hide
  });
  document.addEventListener("mouseout",function(e){
    var ic=e.target.closest&&e.target.closest(".nwi");
    if(ic||box.contains(e.target))scheduleHide();
  });
  document.addEventListener("click",function(e){
    var ic=e.target.closest&&e.target.closest(".nwi");
    if(ic){showFor(ic);pinned=!pinned;ic.classList.toggle("on",pinned);e.stopPropagation();return;}
    if(!box.contains(e.target)){pinned=false;box.classList.remove("show");document.querySelectorAll(".nwi.on").forEach(function(n){n.classList.remove("on");});}
  });
})();

/* Note-icon hover popup — same pattern as news above, but the text comes
   straight off the icon's data-note attribute since notes are per-row
   (per watchlist entry / holding / order), not shared per-symbol like news. */
(function(){
  var box=document.getElementById("noBox"),pinned=false,hideT=null;
  function place(ic){
    var r=ic.getBoundingClientRect(),bw=box.offsetWidth||240;
    box.style.left=Math.max(8,Math.min(r.left,window.innerWidth-bw-8))+"px";
    var top=r.top-box.offsetHeight-8;
    box.style.top=(top<8?r.bottom+8:top)+"px";
  }
  function showFor(ic){
    var note=ic.getAttribute("data-note");if(!note)return;
    clearTimeout(hideT);
    box.textContent=note;box.classList.add("show");place(ic);
  }
  function scheduleHide(){
    if(pinned)return;
    clearTimeout(hideT);
    hideT=setTimeout(function(){if(!pinned)box.classList.remove("show");},300);
  }
  document.addEventListener("mouseover",function(e){
    var ic=e.target.closest&&e.target.closest(".noi");
    if(ic){showFor(ic);return;}
    if(box.contains(e.target))clearTimeout(hideT);
  });
  document.addEventListener("mouseout",function(e){
    var ic=e.target.closest&&e.target.closest(".noi");
    if(ic||box.contains(e.target))scheduleHide();
  });
  document.addEventListener("click",function(e){
    var ic=e.target.closest&&e.target.closest(".noi");
    if(ic){showFor(ic);pinned=!pinned;ic.classList.toggle("on",pinned);e.stopPropagation();return;}
    if(!box.contains(e.target)){pinned=false;box.classList.remove("show");document.querySelectorAll(".noi.on").forEach(function(n){n.classList.remove("on");});}
  });
})();

/* ══ THEME / CLOCK ══ */
var _theme="dark";
function toggleTheme(){_theme=_theme==="dark"?"light":"dark";document.documentElement.setAttribute("data-theme",_theme);var ic=_theme==="dark"?"&#9681;":"&#9788;";document.getElementById("thI").innerHTML=ic;document.getElementById("sbTI").innerHTML=ic;}
function tick(){document.getElementById("clk").textContent=new Date().toLocaleTimeString("en-IN",{hour12:false});var n=new Date(),dow=n.getDay(),t=n.getHours()*100+n.getMinutes(),open=dow>=1&&dow<=5&&t>=915&&t<=1530;var mt=document.getElementById("mktTag");if(mt)mt.style.display=open?"none":"inline";}
setInterval(tick,10000);tick();

/* ══ FEEDBACK ══ */
function openFB(){document.getElementById("fbM").classList.add("open");}
function closeFB(){document.getElementById("fbM").classList.remove("open");["fbEmail","fbSub","fbMsg"].forEach(function(id){document.getElementById(id).value="";});document.querySelectorAll(".tybtn").forEach(function(b,i){b.classList.toggle("sel",i===0);});_fbType="Bug";}
function selT(b){document.querySelectorAll(".tybtn").forEach(function(x){x.classList.remove("sel");});b.classList.add("sel");_fbType=b.textContent.trim();}
function sendFB(){var msg=document.getElementById("fbMsg").value.trim();if(!msg){toast("Please enter a message");return;}var btn=document.getElementById("fbSendBtn");btn.disabled=true;btn.textContent="Sending...";
  google.script.run.withSuccessHandler(function(){closeFB();toast("Feedback sent!");btn.disabled=false;btn.textContent="Send";}).withFailureHandler(function(){toast("Failed — run testFeedbackEmail() from IDE first");btn.disabled=false;btn.textContent="Send";}).sendFeedbackEmailFromWeb(_fbType,document.getElementById("fbSub").value.trim()||"(none)",msg,document.getElementById("fbEmail").value.trim()||"");}
document.getElementById("fbM").addEventListener("click",function(e){if(e.target===this)closeFB();});
document.getElementById("profM").addEventListener("click",function(e){if(e.target===this)closeProfile();});
document.getElementById("trafficM").addEventListener("click",function(e){if(e.target===this)closeTraffic();});

/* ══ DATE / TIME ══ */
function todayStr(){var d=new Date();return d.getFullYear()+"-"+(("0"+(d.getMonth()+1)).slice(-2))+"-"+(("0"+d.getDate()).slice(-2));}
function isToday(s){return s===todayStr();}
function fmtDL(s){if(!s||isToday(s))return"Today";try{return new Date(s+"T00:00:00").toLocaleDateString("en-IN",{day:"2-digit",month:"short"});}catch(e){return s;}}
function isAfterMkt(){var n=new Date(),dow=n.getDay();if(dow===0||dow===6)return true;return n.getHours()*100+n.getMinutes()>1530;}
function shiftDate(dir){var c=document.getElementById("gDate").value||todayStr();var d=new Date(c+"T00:00:00");d.setDate(d.getDate()+dir);while(d.getDay()===0||d.getDay()===6)d.setDate(d.getDate()+dir);var s=d.getFullYear()+"-"+(("0"+(d.getMonth()+1)).slice(-2))+"-"+(("0"+d.getDate()).slice(-2));document.getElementById("gDate").value=s;onDateChange(s);}
function gotoToday(){var td=todayStr();document.getElementById("gDate").value=td;onDateChange(td);}
function onDateChange(dt){_gDate=dt||todayStr();_PL.intraday=null;updSigRng();clearTimeout(_timer);if(_tab==="intraday")fetchIntraday();}
function onTimeChange(t){_asOfTime=t||"";updSigRng();clearTimeout(_timer);if(_tab==="intraday"){_PL.intraday=null;fetchIntraday();}}

/* ══ PER-TAB PERIOD SETTINGS ══ */
function saveTabCfg(){_tabCfg[_tab]={mode:_mode,val:_val,pm:_pm};}
function restoreTabCfg(tab){var cfg=_tabCfg[tab]||{mode:"DAYS",val:1,pm:"ltp"};_mode=cfg.mode;_val=cfg.val;_pm=cfg.pm;
  document.querySelectorAll(".mbtn").forEach(function(b){b.classList.toggle("active",b.dataset.m===_mode);});
  document.getElementById("pval").value=_val;
  document.querySelectorAll(".ptb").forEach(function(b){b.classList.toggle("active",b.dataset.pm===_pm);});
}
function setPM(m){_pm=m;document.querySelectorAll(".ptb").forEach(function(b){b.classList.toggle("active",b.dataset.pm===m);});saveTabCfg();
  if(_tab==="shortterm"&&_st.length)renderST(_st);
  else if(_tab==="intraday"&&_iRows.length)renderIntraday(_iRows);
  else if(_tab==="longterm"&&_ltRows.length)renderLT(_ltRows);
}
/* ST price helpers */
function stP(r){return _pm==="high"&&r.dayHigh?r.dayHigh:r.ltp;}
function stR(r){var p=stP(r);return(p!=null&&r.entry)?(p-r.entry)/r.entry:null;}
function stPr(r){var p=stP(r);return p!=null&&p>r.entry;}
/* Target/SL-hit always use DAY_HIGH/DAY_LOW — independent of the LTP/High
   display toggle (_pm above), which only affects displayed price/return. */
function stTg(r){var h=r.dayHigh!=null?r.dayHigh:r.ltp;if(h==null)return false;var t=r.sheetTarget>0?r.sheetTarget:r.target;return t>0&&h>=t;}
function stSl(r){var lo=r.dayLow!=null?r.dayLow:r.ltp;if(lo==null)return false;var s=r.sheetSL>0?r.sheetSL:r.sl;return s>0&&lo<=s;}
/* Intraday price helpers */
function iP(r){if(_pm==="high")return r.isSell?(r.dayLow||r.ltp):(r.dayHigh||r.ltp);return r.ltp;}
function iR(r){var p=iP(r);return(p!=null&&r.entry)?(r.isSell?(r.entry-p)/r.entry:(p-r.entry)/r.entry):null;}
function iPr(r){var p=iP(r);return p!=null&&(r.isSell?p<r.entry:p>r.entry);}
/* Trust the backend-computed flags (already High/Low based via the IPA sheet
   formulas — see ipaRowToObj) instead of recomputing off the toggled price. */
function iTg(r){return !!r.targetHit;}
function iSlHit(r){return !!r.slHit;}
function ltTg(r){var h=r.dayHigh!=null?r.dayHigh:r.ltp;if(h==null||!r.target)return false;return h>=r.target;}
function ltSl(r){var lo=r.dayLow!=null?r.dayLow:r.ltp;if(lo==null||!r.sl)return false;return lo<=r.sl;}
/* LT helpers — ltP respects the LTP/High toggle (_pm) for DISPLAY only;
   hit-detection above always uses High/Low regardless of this toggle. */
function ltP(r){return _pm==="high"&&r.dayHigh?r.dayHigh:r.ltp;}
function ltR(r){var p=ltP(r);return(p!=null&&r.entry)?(p-r.entry)/r.entry:null;}

/* Period bar events */
var TCFG={
  home:     {modes:["DAYS","WEEKS","MONTHS","YEARS"],kl:["Total","In Profit","In Loss","Target Hits","SL Hits","Exited"],ks:["","","","","",""],kc:["var(--bl)","var(--gr)","var(--rd)","var(--gr)","var(--rd)","var(--pu)"]},
  intraday: {modes:["DAYS","WEEKS","MONTHS","YEARS"],kl:["BUY Active","Short-sell","In Profit","In Loss","Target Hits","SL Hits"],ks:["buy signals","sell/short-sell","in favour","against","target reached","SL Hit"],kc:["var(--gr)","var(--rd)","var(--gr)","var(--rd)","var(--gr)","var(--rd)"]},
  shortterm:{modes:["DAYS","WEEKS","MONTHS","YEARS"],kl:["Total Signals","In Profit","In Loss","Target Hits","SL Hits","Exited"],ks:["unique cycles","price > entry","price <= entry","by day high","by day low","left source"],kc:["var(--bl)","var(--gr)","var(--rd)","var(--gr)","var(--rd)","var(--pu)"]},
  longterm: {modes:["WEEKS","MONTHS","YEARS"],kl:["Total Holdings","Profitable","In Loss","Target Hits","SL Hits","Exited"],ks:["tracked","LTP > entry","LTP <= entry","day-high >= target","LTP <= SL","left source"],kc:["var(--bl)","var(--gr)","var(--rd)","var(--gr)","var(--rd)","var(--pu)"]}
};
function updSigRng(){
  if(_tab==="home"){document.getElementById("srng").textContent="All Time";return;}
  if(_tab==="intraday"&&_mode==="DAYS"&&_val===1){var lbl="Session: "+fmtDL(_gDate);if(_asOfTime)lbl+=" @ "+_asOfTime;document.getElementById("srng").textContent=lbl;return;}
  var lm={DAYS:"Day",WEEKS:"Week",MONTHS:"Month",YEARS:"Year"};
  document.getElementById("srng").textContent="Last "+_val+" "+(_val>1?lm[_mode]+"s":lm[_mode]);
}
document.getElementById("mgrp").addEventListener("click",function(e){
  var btn=e.target.closest(".mbtn");if(!btn||btn.disabled)return;
  document.querySelectorAll(".mbtn").forEach(function(b){b.classList.remove("active");});btn.classList.add("active");
  _mode=btn.dataset.m;document.getElementById("asOfTime").disabled=!(_mode==="DAYS");
  updSigRng();saveTabCfg();clearTimeout(_timer);_PL.st=null;_PL.intraday=null;_PL.lt=null;
  if(_tab==="shortterm")silentFetch();else if(_tab==="intraday")fetchIntraday();else if(_tab==="longterm")fetchLT();
});
document.getElementById("pval").addEventListener("change",function(){
  var v=Math.max(1,parseInt(this.value)||1);_val=v;
  document.getElementById("asOfTime").disabled=!(_mode==="DAYS"&&v===1);
  updSigRng();saveTabCfg();clearTimeout(_timer);_PL.st=null;_PL.intraday=null;_PL.lt=null;
  if(_tab==="shortterm")silentFetch();else if(_tab==="intraday")fetchIntraday();else if(_tab==="longterm")fetchLT();
});

/* ══ TAB SWITCH ══ */
function switchTab(tab){
  saveTabCfg();_tab=tab;restoreTabCfg(tab);
  var cfg=TCFG[tab];
  document.querySelectorAll(".tbtn").forEach(function(b){b.classList.toggle("active",b.dataset.tab===tab);});
  cfg.kl.forEach(function(l,i){var e=document.getElementById("k"+i+"l");if(e)e.textContent=l;});
  cfg.ks.forEach(function(s,i){var e=document.getElementById("k"+i+"s");if(e)e.textContent=s;});
  if(cfg.kc)cfg.kc.forEach(function(c,i){var card=document.querySelector('.kpi[data-ki="'+i+'"]');if(card)card.style.setProperty("--ka",c);});
  document.querySelectorAll(".mbtn").forEach(function(b){b.disabled=cfg.modes.indexOf(b.dataset.m)===-1;});
  var dn=document.getElementById("dnavGrp"),dnb=document.getElementById("dnBsp");
  dn.classList.toggle("on",tab==="intraday");if(dnb)dnb.style.display=tab==="intraday"?"block":"none";
  document.getElementById("livLbl").textContent={home:"All Time",intraday:"Session",shortterm:"Live",longterm:"Live"}[tab]||"Live";
  updSigRng();
  ["cnt-active","cnt-home","cnt-intraday","cnt-longterm"].forEach(function(id){document.getElementById(id).style.display="none";});
  clrKPIs();clearTimeout(_timer);closeDrST();closeDrI();closeDrLT();
  if(tab==="home"){document.getElementById("cnt-home").style.display="block";clearFooter();hideLoader();}
  else if(tab==="intraday"){document.getElementById("cnt-intraday").style.display="block";if(_PL.intraday&&isToday(_gDate)&&!_asOfTime)renderIntraday(_PL.intraday);else fetchIntraday();}
  else if(tab==="shortterm"){document.getElementById("cnt-active").style.display="block";if(_PL.st)renderST(_PL.st);else silentFetch();}
  else if(tab==="longterm"){document.getElementById("cnt-longterm").style.display="block";clearFooter();if(_PL.lt)renderLT(_PL.lt);else fetchLT();}
  schedRef();
}

/* ══ RENDER — no opacity changes, instant update ══ */
function renderST(rows){
  _st=rows;_PL.st=rows;
  var withP=rows.filter(function(r){return r.ltp!=null;});
  var pro=withP.filter(function(r){return stPr(r);}).length,los=withP.length-pro;
  var tgt=rows.filter(function(r){return stTg(r);}).length,sl=rows.filter(function(r){return stSl(r);}).length;
  var ex=rows.filter(function(r){return r.exitFlag==="EXITED";}).length,run=rows.filter(function(r){return r.exitFlag!=="EXITED";}).length;
  var rate=withP.length?pro/withP.length:0;
  if(_tab==="shortterm"){setK(0,rows.length,rows.length);setK(1,pro,rows.length);setK(2,los,rows.length);setK(3,tgt,rows.length);setK(4,sl,rows.length);setK(5,ex,rows.length);}
  animD("darc","dpct",rate,null);var srb=document.getElementById("srBdg");if(srb)srb.textContent=Math.round(rate*100)+"%";
  var es={lgP:pro,lgL:los,lgR:run,lgE:ex};Object.keys(es).forEach(function(k){var e=document.getElementById(k);if(e)e.textContent=es[k];});
  var s=withP.slice().sort(function(a,b){return stR(b)-stR(a);});
  /* ST performer row: #, Symbol, Entry, SL, LTP, High, Return, Status */
  var mkRow=function(r,i){var ret=stR(r);var slV=r.sl||r.sheetSL;return'<tr><td style="color:var(--mu)">'+(i+1)+'</td>'+symCell(r.symbol,r.name||"",r._consolidated||"","shortterm")+'<td style="color:var(--dm)">'+fNZ(r.entry)+'</td><td>'+fNZ(r.ltp)+'</td><td style="color:var(--rd)">'+fNZ(slV)+'</td><td>'+hC(r.dayHigh)+'</td><td class="'+(ret!=null&&ret>=0?"rp":"rn")+'">'+fR(ret)+'</td><td>'+stPill(r)+'</td></tr>';};
  upd("topBdy",s.slice(0,12).map(mkRow).join("")||'<tr class="emr"><td colspan="8">No data</td></tr>');
  upd("wstBdy",s.slice().sort(function(a,b){return stR(a)-stR(b);}).slice(0,12).map(mkRow).join("")||'<tr class="emr"><td colspan="8">No data</td></tr>');
  renderBars("bsec",[{l:"Total",v:rows.length,c:"var(--bl)"},{l:"Profit",v:pro,c:"var(--gr)"},{l:"Loss",v:los,c:"var(--rd)"},{l:"Tgt",v:tgt,c:"var(--gr)"},{l:"SL",v:sl,c:"var(--rd)"},{l:"Exited",v:ex,c:"var(--pu)"}]);
  updFoot(rows,stR);if(_stF)refreshStDr();
}
function renderIntraday(rows){
  _iRows=rows;if(isToday(_gDate)&&!_asOfTime)_PL.intraday=rows;
  var eff=isAfterMkt()?rows.map(function(r){return r.exitFlag!=="EXITED"?Object.assign({},r,{exitFlag:"EXITED"}):r;}):rows;
  var withP=eff.filter(function(r){return r.ltp!=null;});
  var pro=withP.filter(function(r){return iPr(r);}).length,los=withP.length-pro;
  var buy=eff.filter(function(r){return!r.isSell;}).length,sell=eff.filter(function(r){return r.isSell;}).length;
  var tgt=eff.filter(function(r){return iTg(r);}).length;
  var slh=eff.filter(function(r){return iSlHit(r);}).length;
  var rate=withP.length?pro/withP.length:0;
  if(_tab==="intraday"){setK(0,buy,eff.length);setK(1,sell,eff.length);setK(2,pro,eff.length);setK(3,los,eff.length);setK(4,tgt,eff.length);setK(5,slh,eff.length);}
  animD("iarc","idpct",rate,"var(--cy)");var isrb=document.getElementById("isrBdg");if(isrb)isrb.textContent=Math.round(rate*100)+"%";
  [{id:"ilgP",v:pro},{id:"ilgL",v:los},{id:"ilgB",v:buy},{id:"ilgS",v:sell}].forEach(function(x){var e=document.getElementById(x.id);if(e)e.textContent=x.v;});
  var byR=withP.slice().sort(function(a,b){return iR(b)-iR(a);});
  var iMkRow=function(r,i){var ret=iR(r);return'<tr><td style="color:var(--mu)">'+(i+1)+'</td>'+symCell(r.symbol,r.name||"",r._consolidated||"","intraday")+'<td>'+oPill(r.order)+'</td><td style="color:var(--dm)">'+fNZ(r.entry)+'</td><td>'+fNZ(r.ltp)+'</td><td class="'+(ret!=null&&ret>=0?"rp":"rn")+'">'+fR(ret)+'</td><td>'+iPill(r)+'</td></tr>';};
  upd("itopBdy",byR.slice(0,12).map(iMkRow).join("")||'<tr class="emr"><td colspan="7">No data</td></tr>');
  upd("iwstBdy",withP.slice().sort(function(a,b){return iR(a)-iR(b);}).slice(0,12).map(iMkRow).join("")||'<tr class="emr"><td colspan="7">No data</td></tr>');
  renderBars("ibsec",[{l:"BUY",v:buy,c:"var(--gr)"},{l:"Sell",v:sell,c:"var(--rd)"},{l:"Profit",v:pro,c:"var(--gr)"},{l:"Loss",v:los,c:"var(--rd)"},{l:"Tgt",v:tgt,c:"var(--cy)"}]);
  updFoot(eff,iR);if(_iF)refreshIDr();
}
function renderLT(rows){
  _ltRows=rows;_PL.lt=rows;
  var withP=rows.filter(function(r){return r.ltp!=null;});
  var pro=withP.filter(function(r){return ltP(r)>r.entry;}).length,los=withP.length-pro;
  var ex=rows.filter(function(r){return r.exitFlag==="EXITED";}).length,hold=rows.filter(function(r){return r.exitFlag!=="EXITED";}).length;
  var ltTgtHits=withP.filter(function(r){return ltTg(r);}).length;
  var ltSlHits=withP.filter(function(r){return ltSl(r);}).length;
  var rate=withP.length?pro/withP.length:0;
  if(_tab==="longterm"){setK(0,rows.length,rows.length);setK(1,pro,rows.length);setK(2,los,rows.length);setK(3,ltTgtHits,rows.length);setK(4,ltSlHits,rows.length);setK(5,ex,rows.length);}
  animD("ltarc","ltpct",rate,"var(--am)");var ltSrb=document.getElementById("ltSrBdg");if(ltSrb)ltSrb.textContent=Math.round(rate*100)+"%";
  [{id:"ltLP",v:pro},{id:"ltLL",v:los},{id:"ltLH",v:hold},{id:"ltLE",v:ex}].forEach(function(x){var e=document.getElementById(x.id);if(e)e.textContent=x.v;});
  var s=withP.slice().sort(function(a,b){return ltR(b)-ltR(a);});
  var mkR=function(r,i){var ret=ltR(r);return'<tr><td style="color:var(--mu)">'+(i+1)+'</td>'+symCell(r.symbol,r.name||"",r._consolidated||"","longterm")+'<td style="color:var(--dm)">'+fNZ(r.entry)+'</td><td>'+fNZ(r.ltp)+'</td><td>'+hC(r.dayHigh)+'</td><td class="'+(ret!=null&&ret>=0?"rp":"rn")+'">'+fR(ret)+'</td><td>'+ltPill(r)+'</td></tr>';};
  upd("lttopBdy",s.slice(0,12).map(mkR).join("")||'<tr class="emr"><td colspan="7">Run runLTSystem() from Apps Script IDE first.</td></tr>');
  upd("ltwstBdy",s.slice().sort(function(a,b){return ltR(a)-ltR(b);}).slice(0,12).map(mkR).join("")||'<tr class="emr"><td colspan="7">-</td></tr>');
  renderBars("ltbsec",[{l:"Total",v:rows.length,c:"var(--am)"},{l:"Profit",v:pro,c:"var(--gr)"},{l:"Loss",v:los,c:"var(--rd)"},{l:"Exited",v:ex,c:"var(--pu)"}]);
  updFoot(withP,ltR);if(_ltF)refreshDrLT();
}

/* ══ FETCHING — all simultaneous, no delays, no opacity changes ══ */
function silentFetch(){if(_stBusy)return;_stBusy=true;google.script.run.withSuccessHandler(function(r){_stBusy=false;renderST(r.rows||[]);}).withFailureHandler(function(){_stBusy=false;}).runAnalysisFromWeb(_mode,_val);}
function fetchIntraday(){
  var ed=_gDate||todayStr(),single=(_mode==="DAYS"&&_val===1);
  if(single)google.script.run.withSuccessHandler(function(r){_iRows=r.rows||[];renderIntraday(_iRows);schedRef();}).withFailureHandler(function(){}).getIntradayAnalysisFromWeb(ed,_asOfTime||null);
  else google.script.run.withSuccessHandler(function(r){_iRows=r.rows||[];renderIntraday(_iRows);schedRef();}).withFailureHandler(function(){}).getIntradayAnalysisByPeriodFromWeb(_mode,_val,ed);
}
function fetchLT(){if(_ltBusy)return;_ltBusy=true;google.script.run.withSuccessHandler(function(r){_ltBusy=false;renderLT(r.rows||[]);}).withFailureHandler(function(){_ltBusy=false;upd("lttopBdy",'<tr class="emr"><td colspan="7">Run runLTSystem() from Apps Script IDE.</td></tr>');}).getLTAnalysisDataFromWeb(_mode,_val);}
function schedRef(){clearTimeout(_timer);_timer=setTimeout(function(){if(_tab==="shortterm")silentFetch();else if(_tab==="intraday")fetchIntraday();else if(_tab==="longterm")fetchLT();},RFRSH);}
function hideLoader(){var e=document.getElementById("loader");if(!e)return;e.classList.add("gn");setTimeout(function(){if(e.parentNode)e.parentNode.removeChild(e);},600);}
function preloadAll(){
  hideLoader();fetchTicker();
  google.script.run.withSuccessHandler(function(r){_PL.st=r.rows||[];_st=_PL.st;if(_tab==="shortterm")renderST(_st);}).runAnalysisFromWeb("DAYS",1);
  google.script.run.withSuccessHandler(function(r){_PL.intraday=r.rows||[];_iRows=_PL.intraday;if(_tab==="intraday")renderIntraday(_iRows);}).getIntradayAnalysisFromWeb(todayStr(),null);
  google.script.run.withSuccessHandler(function(r){_PL.lt=r.rows||[];if(_tab==="longterm")renderLT(_PL.lt);}).getLTAnalysisDataFromWeb(_tabCfg.longterm.mode,_tabCfg.longterm.val);
  /* Live-panel data is not visible until the user opens that panel, so it's
     deferred slightly instead of firing in the same instant as the 3 calls
     above — those 3 were competing for the same execution window as the
     tab data you actually see on open, which was part of the slowness. */
  setTimeout(function(){
    google.script.run.withSuccessHandler(function(r){_PL.lpID=r.rows||[];}).getLiveSignalsFromWeb("intraday");
    google.script.run.withSuccessHandler(function(r){_PL.lpST=r.rows||[];}).getLiveSignalsFromWeb("shortterm");
    google.script.run.withSuccessHandler(function(r){_PL.lpLT=r.rows||[];}).getLiveSignalsFromWeb("longterm");
  },600);
}

/* ══ DRAWERS ══ */
var STDR={all:{t:"All Signals",c:"#4da6ff",f:function(r){return true;},s:function(a,b){return(new Date(b.signalDate||0))-(new Date(a.signalDate||0));}},profit:{t:"In Profit",c:"var(--gr)",f:function(r){return r.ltp!=null&&stPr(r);},s:function(a,b){return stR(b)-stR(a);}},loss:{t:"In Loss",c:"var(--rd)",f:function(r){return r.ltp!=null&&!stPr(r);},s:function(a,b){return stR(a)-stR(b);}},target:{t:"Target Hit",c:"var(--gr)",f:function(r){return stTg(r);},s:function(a,b){return stR(b)-stR(a);}},sl:{t:"SL Hit",c:"var(--rd)",f:function(r){return stSl(r);},s:function(a,b){return stR(a)-stR(b);}},exited:{t:"Exited",c:"#b48eff",f:function(r){return r.exitFlag==="EXITED";},s:function(a,b){return stR(b)-stR(a);}}};
var IDR={ibuy:{t:"BUY",c:"var(--gr)",f:function(r){return!r.isSell;}},isell:{t:"Short-sell",c:"var(--rd)",f:function(r){return r.isSell;}},iprofit:{t:"In Profit",c:"var(--gr)",f:function(r){return r.ltp!=null&&iPr(r);}},iloss:{t:"In Loss",c:"var(--rd)",f:function(r){return r.ltp!=null&&!iPr(r);}},itarget:{t:"Target Hit",c:"var(--cy)",f:function(r){return iTg(r);}},isl:{t:"SL Hit",c:"var(--rd)",f:function(r){return iSlHit(r);}}};
var LTDR={ltall:{t:"All Holdings",c:"var(--am)",f:function(r){return true;},s:function(a,b){return ltR(b)-ltR(a);}},ltprofit:{t:"Profitable",c:"var(--gr)",f:function(r){return r.ltp!=null&&r.ltp>r.entry;},s:function(a,b){return ltR(b)-ltR(a);}},ltloss:{t:"In Loss",c:"var(--rd)",f:function(r){return r.ltp!=null&&r.ltp<=r.entry;},s:function(a,b){return ltR(a)-ltR(b);}},lttarget:{t:"Target Hit",c:"var(--gr)",f:function(r){return ltTg(r);},s:function(a,b){return ltR(b)-ltR(a);}},ltsl:{t:"SL Hit",c:"var(--rd)",f:function(r){return ltSl(r);},s:function(a,b){return ltR(a)-ltR(b);}},ltexited:{t:"Exited",c:"#b48eff",f:function(r){return r.exitFlag==="EXITED";},s:function(a,b){return ltR(b)-ltR(a);}}};
var _stKeys=["all","profit","loss","target","sl","exited"];
var _iKeys=["ibuy","isell","iprofit","iloss","itarget","isl"];
var _ltKeys=["ltall","ltprofit","ltloss","lttarget","ltsl","ltexited"];
document.querySelectorAll(".kpi[data-ki]").forEach(function(card){card.addEventListener("click",function(){var idx=parseInt(card.dataset.ki);if(_tab==="shortterm"){var key=_stKeys[idx];if(_stF===key){closeDrST();return;}openStDr(key);}else if(_tab==="intraday"){var key=_iKeys[idx];if(_iF===key){closeDrI();return;}openIDr(key);}else if(_tab==="longterm"){var key=_ltKeys[idx];if(!LTDR[key]||!LTDR[key].f)return;if(_ltF===key){closeDrLT();return;}openDrLT(key);}});});
/* ST DRAWER */
var ST_ACC={symbol:function(r){return r.symbol;},since:function(r){return r.signalDate;},entry:function(r){return r.entry;},ltp:function(r){return r.ltp;},sl:function(r){return r.sl||r.sheetSL;},target:function(r){return r.sheetTarget||r.target;},high:function(r){return r.dayHigh;},low:function(r){return r.dayLow;},return:function(r){return stR(r);},thit:function(r){return stTg(r)?1:0;},slhit:function(r){return stSl(r)?1:0;},days:function(r){return r.tradingDays;},cycles:function(r){return r._consolidated||0;},status:function(r){return r.exitFlag==="EXITED"?(stPr(r)?"success":"failed"):"running";},exit:function(r){return r.exitDate;},exited:function(r){return r.exitDate;}};
var I_ACC={symbol:function(r){return r.symbol;},order:function(r){return r.order;},since:function(r){return r.entryDateTime||r.tradeDate;},entry:function(r){return r.entry;},ltp:function(r){return r.ltp;},sl:function(r){return r.sl;},high:function(r){return r.dayHigh;},low:function(r){return r.dayLow;},return:function(r){return iR(r);},target:function(r){return r.target;},thit:function(r){return r.targetHit?1:0;},slhit:function(r){return r.slHit?1:0;},days:function(r){return r.tradingMins;},mins:function(r){return r.tradingMins;},cycles:function(r){return r._consolidated||0;},status:function(r){return r.exitFlag==="EXITED"?(iPr(r)?"success":"failed"):"running";}};
var LT_ACC={symbol:function(r){return r.symbol;},since:function(r){return r.signalDate;},entry:function(r){return r.entry;},ltp:function(r){return r.ltp;},sl:function(r){return r.sl;},target:function(r){return r.target;},high:function(r){return r.dayHigh;},low:function(r){return r.dayLow;},return:function(r){return ltR(r);},thit:function(r){return ltTg(r)?1:0;},slhit:function(r){return ltSl(r)?1:0;},days:function(r){return r.tradingDays;},cycles:function(r){return r._consolidated||0;},status:function(r){return r.exitFlag==="EXITED"?(ltR(r)>0?"success":"failed"):"running";},exit:function(r){return r.exitDate;},exited:function(r){return r.exitDate;}};

function openStDr(key){var cfg=STDR[key];if(!cfg)return;_stF=key;_stSort={key:null,dir:1};_stDR=_st.filter(cfg.f).sort(cfg.s);var dr=document.getElementById("dr");dr.style.setProperty("--drc",cfg.c);document.getElementById("drttl").textContent=cfg.t;document.getElementById("drsc").value="";document.querySelectorAll("#drThead th").forEach(function(th){th.classList.remove("srt-asc","srt-desc");});document.querySelectorAll(".kpi").forEach(function(c){c.classList.remove("af");});var idx=_stKeys.indexOf(key);if(idx>=0){var tile=document.querySelector('.kpi[data-ki="'+idx+'"]');if(tile)tile.classList.add("af");}renderStDrRows(_stDR);document.getElementById("drcnt").textContent=_stDR.length+" signals";dr.classList.add("open");setTimeout(function(){dr.scrollIntoView({behavior:"smooth",block:"nearest"});},60);}
function renderStDrRows(rows){var tb=document.getElementById("drbdy");if(!rows.length){tb.innerHTML='<tr><td colspan="15" style="text-align:center;padding:20px;color:var(--dm)">None</td></tr>';return;}
tb.innerHTML=rows.map(function(r,i){var ret=stR(r),h=r.ltp!=null,rc=!h?"":ret>=0?"rp":"rn";var tHit=stTg(r)?'<span style="color:var(--gr)">&#10003;</span>':'-';var slHit=stSl(r)?'<span style="color:var(--rd)">&#10003;</span>':'-';var tgtVal=r.sheetTarget||r.target;var slVal=r.sl||r.sheetSL;var eon=r.exitFlag==="EXITED"&&r.exitDate?'<span style="color:var(--pu)">'+r.exitDate+'</span>':'-';return'<tr><td style="color:var(--mu)">'+(i+1)+'</td>'+symCell(r.symbol,r.name||"",r._consolidated||"","shortterm")+'<td>'+(r.signalDate||"-")+'</td><td>'+fNZ(r.entry)+'</td><td>'+(h?fNZ(r.ltp):'-')+'</td><td style="color:var(--rd)">'+fNZ(slVal)+'</td><td style="color:var(--gr)">'+fNZ(tgtVal)+'</td><td>'+hC(r.dayHigh)+'</td><td>'+(r.dayLow!=null?'<span style="color:var(--cy)">'+fNZ(r.dayLow)+'</span>':'-')+'</td><td class="'+rc+'">'+(h?fR(ret):'-')+'</td><td>'+tHit+'</td><td>'+slHit+'</td><td>'+(r.tradingDays||"-")+'</td><td>'+stPill(r)+'</td><td>'+eon+'</td></tr>';}).join("");}
function refreshStDr(){if(!_stF)return;var cfg=STDR[_stF];if(!cfg)return;_stDR=_st.filter(cfg.f).sort(cfg.s);var q=document.getElementById("drsc").value;var out=q?_stDR.filter(function(r){return kpiRowMatches_(r,q,ST_ACC);}):_stDR;out=kpiSortRows_(out,_stSort,ST_ACC);renderStDrRows(out);document.getElementById("drcnt").textContent=_stDR.length+" signals";}
function closeDrST(){document.getElementById("dr").classList.remove("open");document.querySelectorAll(".kpi").forEach(function(c){c.classList.remove("af");});_stF=null;document.getElementById("drsc").value="";}
document.getElementById("drsc").addEventListener("input",function(){var q=this.value;var out=q?_stDR.filter(function(r){return kpiRowMatches_(r,q,ST_ACC);}):_stDR;renderStDrRows(kpiSortRows_(out,_stSort,ST_ACC));});
function sortDr(key){kpiSortClick_("drThead",key,_stSort,function(){var q=document.getElementById("drsc").value;var out=q?_stDR.filter(function(r){return kpiRowMatches_(r,q,ST_ACC);}):_stDR;renderStDrRows(kpiSortRows_(out,_stSort,ST_ACC));});}
/* INTRADAY DRAWER */
function openIDr(key){var cfg=IDR[key];if(!cfg)return;_iF=key;_iSort={key:null,dir:1};_iDR=_iRows.filter(cfg.f).sort(function(a,b){return iR(b)-iR(a);});var dr=document.getElementById("idr");dr.style.setProperty("--drc",cfg.c);document.getElementById("idrttl").textContent=cfg.t;document.getElementById("idrsc").value="";document.querySelectorAll("#idrThead th").forEach(function(th){th.classList.remove("srt-asc","srt-desc");});document.querySelectorAll(".kpi").forEach(function(c){c.classList.remove("af");});var idx=_iKeys.indexOf(key);if(idx>=0){var tile=document.querySelector('.kpi[data-ki="'+idx+'"]');if(tile)tile.classList.add("af");}renderIDrRows(_iDR);document.getElementById("idrcnt").textContent=_iDR.length+" signals";dr.classList.add("open");setTimeout(function(){dr.scrollIntoView({behavior:"smooth",block:"nearest"});},60);}
function renderIDrRows(rows){var tb=document.getElementById("idrbdy");if(!rows.length){tb.innerHTML='<tr><td colspan="15" style="text-align:center;padding:20px;color:var(--dm)">None</td></tr>';return;}
tb.innerHTML=rows.map(function(r,i){var ret=iR(r),rc=ret==null?"":ret>=0?"rp":"rn";return'<tr><td style="color:var(--mu)">'+(i+1)+'</td>'+symCell(r.symbol,r.name||"",r._consolidated||"","intraday")+'<td>'+oPill(r.order)+'</td><td style="color:var(--dm);font-size:10px">'+(r.entryDateTime||r.tradeDate||"-")+'</td><td>'+fNZ(r.entry)+'</td><td>'+fNZ(r.ltp)+'</td><td style="color:var(--rd)">'+fNZ(r.sl)+'</td><td>'+hC(r.dayHigh)+'</td><td>'+(r.dayLow!=null?'<span style="color:var(--cy)">'+fNZ(r.dayLow)+'</span>':'-')+'</td><td class="'+rc+'">'+fR(ret)+'</td><td>'+fNZ(r.target)+'</td><td>'+(r.targetHit?'<span style="color:var(--gr)">&#10003;</span>':'-')+'</td><td>'+(r.slHit?'<span style="color:var(--rd)">&#10003;</span>':'-')+'</td><td>'+r.tradingMins+'m</td><td>'+iPill(r)+'</td></tr>';}).join("");}
function refreshIDr(){if(!_iF)return;var cfg=IDR[_iF];if(!cfg)return;_iDR=_iRows.filter(cfg.f).sort(function(a,b){return iR(b)-iR(a);});var q=document.getElementById("idrsc").value;var out=q?_iDR.filter(function(r){return kpiRowMatches_(r,q,I_ACC);}):_iDR;out=kpiSortRows_(out,_iSort,I_ACC);renderIDrRows(out);document.getElementById("idrcnt").textContent=_iDR.length+" signals";}
function closeDrI(){document.getElementById("idr").classList.remove("open");document.querySelectorAll(".kpi").forEach(function(c){c.classList.remove("af");});_iF=null;document.getElementById("idrsc").value="";}
document.getElementById("idrsc").addEventListener("input",function(){var q=this.value;var out=q?_iDR.filter(function(r){return kpiRowMatches_(r,q,I_ACC);}):_iDR;renderIDrRows(kpiSortRows_(out,_iSort,I_ACC));});
function sortIDr(key){kpiSortClick_("idrThead",key,_iSort,function(){var q=document.getElementById("idrsc").value;var out=q?_iDR.filter(function(r){return kpiRowMatches_(r,q,I_ACC);}):_iDR;renderIDrRows(kpiSortRows_(out,_iSort,I_ACC));});}
/* LT DRAWER */
function openDrLT(key){var cfg=LTDR[key];if(!cfg||!cfg.f)return;_ltF=key;_ltSort={key:null,dir:1};_ltDR=_ltRows.filter(cfg.f).sort(cfg.s||function(a,b){return ltR(b)-ltR(a);});var dr=document.getElementById("ltdr");dr.style.setProperty("--drc",cfg.c);document.getElementById("ltdrttl").textContent=cfg.t;document.getElementById("ltdrsc").value="";document.querySelectorAll("#ltdrThead th").forEach(function(th){th.classList.remove("srt-asc","srt-desc");});document.querySelectorAll(".kpi").forEach(function(c){c.classList.remove("af");});var idx=_ltKeys.indexOf(key);if(idx>=0){var tile=document.querySelector('.kpi[data-ki="'+idx+'"]');if(tile)tile.classList.add("af");}renderLTDrRows(_ltDR);document.getElementById("ltdrcnt").textContent=_ltDR.length+" holdings";dr.classList.add("open");setTimeout(function(){dr.scrollIntoView({behavior:"smooth",block:"nearest"});},60);}
function renderLTDrRows(rows){var tb=document.getElementById("ltdrbdy");if(!rows.length){tb.innerHTML='<tr><td colspan="15" style="text-align:center;padding:20px;color:var(--dm)">None</td></tr>';return;}
tb.innerHTML=rows.map(function(r,i){var ret=ltR(r),rc=ret!=null&&ret>=0?"rp":"rn";var eon=r.exitFlag==="EXITED"&&r.exitDate?'<span style="color:var(--pu)">'+r.exitDate+'</span>':'-';return'<tr><td style="color:var(--mu)">'+(i+1)+'</td>'+symCell(r.symbol,r.name||"",r._consolidated||"","longterm")+'<td>'+(r.signalDate||"-")+'</td><td>'+fNZ(r.entry)+'</td><td>'+fNZ(r.ltp)+'</td><td style="color:var(--rd)">'+fNZ(r.sl)+'</td><td style="color:var(--gr)">'+fNZ(r.target)+'</td><td>'+hC(r.dayHigh)+'</td><td>'+(r.dayLow!=null?'<span style="color:var(--cy)">'+fNZ(r.dayLow)+'</span>':'-')+'</td><td class="'+rc+'">'+fR(ret)+'</td><td>'+(r.targetHit?'<span style="color:var(--gr)">&#10003;</span>':'-')+'</td><td>'+(r.slHit?'<span style="color:var(--rd)">&#10003;</span>':'-')+'</td><td>'+(r.tradingDays||"-")+'</td><td>'+ltPill(r)+'</td><td>'+eon+'</td></tr>';}).join("");}
function refreshDrLT(){if(!_ltF)return;var cfg=LTDR[_ltF];if(!cfg||!cfg.f)return;_ltDR=_ltRows.filter(cfg.f).sort(cfg.s||function(a,b){return ltR(b)-ltR(a);});var q=document.getElementById("ltdrsc").value;var out=q?_ltDR.filter(function(r){return kpiRowMatches_(r,q,LT_ACC);}):_ltDR;renderLTDrRows(kpiSortRows_(out,_ltSort,LT_ACC));}
function closeDrLT(){document.getElementById("ltdr").classList.remove("open");document.querySelectorAll(".kpi").forEach(function(c){c.classList.remove("af");});_ltF=null;document.getElementById("ltdrsc").value="";}
document.getElementById("ltdrsc").addEventListener("input",function(){var q=this.value;var out=q?_ltDR.filter(function(r){return kpiRowMatches_(r,q,LT_ACC);}):_ltDR;renderLTDrRows(kpiSortRows_(out,_ltSort,LT_ACC));});
function sortLtDr(key){kpiSortClick_("ltdrThead",key,_ltSort,function(){var q=document.getElementById("ltdrsc").value;var out=q?_ltDR.filter(function(r){return kpiRowMatches_(r,q,LT_ACC);}):_ltDR;renderLTDrRows(kpiSortRows_(out,_ltSort,LT_ACC));});}


document.addEventListener("keydown",function(e){if(e.key==="Escape"){closeFB();closeLP();closeAI();closeDrST();closeDrI();closeDrLT();closeProfile();closeTraffic();}});

/* Block copying signal data out of the dashboard. Inputs/textareas are
   exempted so typing/pasting your own text (AI question, feedback form,
   period value) still works normally. This is a deterrent, not a hard
   guarantee — anyone using browser dev tools or view-source can still
   reach the underlying data; there's no fully bulletproof way to prevent
   that from a web page. */
function _inField(e){return e.target.closest&&e.target.closest("input,textarea");}
document.addEventListener("copy",function(e){if(!_inField(e))e.preventDefault();});
document.addEventListener("cut",function(e){if(!_inField(e))e.preventDefault();});
document.addEventListener("contextmenu",function(e){if(!_inField(e))e.preventDefault();});

/* ══ AUTH ══ */
var _authToken=null,_pendingEmail="",_myReferralCode="";
var _referralCode=(function(){try{return new URLSearchParams(window.location.search).get("ref")||"";}catch(e){return "";}})();
function showAuthGate(){
  document.getElementById("authGate").classList.remove("hide");hideLoader();
  try{
    var reason=sessionStorage.getItem("sdLogoutReason");
    if(reason){showErr("loginErr",reason);document.getElementById("loginErr").style.background="rgba(255,180,60,.12)";document.getElementById("loginErr").style.color="var(--am)";sessionStorage.removeItem("sdLogoutReason");}
  }catch(e){}
}
function hideAuthGate(){document.getElementById("authGate").classList.add("hide");}
function showAuthScreen(name){
  var cur=document.querySelector(".authScreen.on");
  document.querySelectorAll(".authScreen").forEach(function(s){s.classList.remove("on");});
  document.getElementById("auth"+name.charAt(0).toUpperCase()+name.slice(1)).classList.add("on");
  document.querySelectorAll(".authTab").forEach(function(t,i){t.classList.toggle("sel",(name==="login"&&i===0)||(name==="signup"&&i===1));});
  // Only clear errors when actually changing screens, not staying on same
  var prev=cur?cur.id:"";
  var next="auth"+name.charAt(0).toUpperCase()+name.slice(1);
  if(prev!==next){["loginErr","signupErr","verifyErr","forgotErr","resetErr","mfaErr"].forEach(function(id){var e=document.getElementById(id);if(e)e.classList.remove("show");});}
}
function showErr(id,msg){var e=document.getElementById(id);e.textContent=msg;e.classList.add("show");}
function afterAuthSuccess(token,displayName){
  _authToken=token;localStorage.setItem("sdToken",token);
  hideAuthGate();initDashboardApp();
}
function doLogin(){
  var email=document.getElementById("loginEmail").value.trim();
  var pass=document.getElementById("loginPass").value;
  var btn=document.getElementById("loginBtn");
  document.getElementById("loginErr").classList.remove("show");
  if(!email||!pass){showErr("loginErr","Enter your email and password.");return;}
  btn.disabled=true;
  google.script.run.withSuccessHandler(function(r){
    btn.disabled=false;
    if(r&&r.success&&r.mfaRequired){_mfaToken=r.mfaToken;document.getElementById("mfaCode").value="";showAuthScreen("mfa");setTimeout(function(){document.getElementById("mfaCode").focus();},50);} // v2: two-step verification
    else if(r&&r.success){afterAuthSuccess(r.token,r.displayName);}
    else if(r&&r.needsVerification){_pendingEmail=email;showAuthScreen("verify");}
    else{showErr("loginErr",(r&&r.error)||"Login failed.");}
  }).withFailureHandler(function(e){btn.disabled=false;showErr("loginErr","Server error: "+(e&&e.message?e.message:""));}).loginFromWeb(email,pass);
}
var _mfaToken="";
function doMfaVerify(){
  var code=document.getElementById("mfaCode").value.trim();
  var btn=document.getElementById("mfaBtn");
  document.getElementById("mfaErr").classList.remove("show");
  if(!code){showErr("mfaErr","Enter the 6-digit code from your authenticator app (or a recovery code).");return;}
  btn.disabled=true;
  google.script.run.withSuccessHandler(function(r){
    btn.disabled=false;
    if(r&&r.success){
      if(r.usedRecovery){try{sessionStorage.setItem("sdMfaNote","You signed in with a recovery code ("+r.recoveryLeft+" left). Generate new ones under Settings \u2192 Security & Account.");}catch(e){}}
      afterAuthSuccess(r.token,r.displayName);
      try{var n=sessionStorage.getItem("sdMfaNote");if(n){sessionStorage.removeItem("sdMfaNote");setTimeout(function(){alert(n);},600);}}catch(e){}
    }
    else if(r&&r.mfaExpired){showAuthScreen("login");showErr("loginErr",(r&&r.error)||"Your sign-in timed out. Please log in again.");}
    else{showErr("mfaErr",(r&&r.error)||"Verification failed.");document.getElementById("mfaCode").select();}
  }).withFailureHandler(function(e){btn.disabled=false;showErr("mfaErr","Server error: "+(e&&e.message?e.message:""));}).mfaVerifyLoginFromWeb(_mfaToken,code);
}
function mfaBackToLogin(){_mfaToken="";document.getElementById("loginPass").value="";showAuthScreen("login");}
function doSignup(){
  var name=document.getElementById("suName").value.trim();
  var email=document.getElementById("suEmail").value.trim();
  var pass=document.getElementById("suPass").value;
  var pass2=document.getElementById("suPass2").value;
  var agree=document.getElementById("suAgree").checked;
  var btn=document.getElementById("signupBtn");
  document.getElementById("signupErr").classList.remove("show");
  if(!name||!email||!pass){showErr("signupErr","Please fill in all fields.");return;}
  if(pass!==pass2){showErr("signupErr","Passwords do not match.");return;}
  if(!agree){showErr("signupErr","Please agree to the Privacy Policy to continue.");return;}
  btn.disabled=true;
  google.script.run.withSuccessHandler(function(r){
    btn.disabled=false;
    if(r&&r.success){_pendingEmail=email;showAuthScreen("verify");}
    else{showErr("signupErr",(r&&r.error)||"Signup failed.");}
  }).withFailureHandler(function(e){btn.disabled=false;showErr("signupErr","Server error: "+(e&&e.message?e.message:""));}).signupFromWeb(name,email,pass,pass2,agree===true,_referralCode);
}
function doVerify(){
  var otp=document.getElementById("verifyOtp").value.trim();
  var btn=document.getElementById("verifyBtn");
  document.getElementById("verifyErr").classList.remove("show");
  if(!otp){showErr("verifyErr","Enter the code sent to your email.");return;}
  btn.disabled=true;
  google.script.run.withSuccessHandler(function(r){
    btn.disabled=false;
    if(r&&r.success){afterAuthSuccess(r.token,r.displayName);}
    else{showErr("verifyErr",(r&&r.error)||"Verification failed.");}
  }).withFailureHandler(function(e){btn.disabled=false;showErr("verifyErr","Server error: "+(e&&e.message?e.message:""));}).verifyEmailFromWeb(_pendingEmail,otp);
}
function doResend(){
  document.getElementById("verifyErr").classList.remove("show");
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.success)showErr("verifyErr","A new code has been sent.");
    else showErr("verifyErr",(r&&r.error)||"Could not resend code.");
  }).withFailureHandler(function(){}).resendVerificationFromWeb(_pendingEmail);
}
function doForgot(){
  var email=document.getElementById("forgotEmail").value.trim();
  var btn=document.getElementById("forgotBtn");
  document.getElementById("forgotErr").classList.remove("show");
  if(!email){showErr("forgotErr","Enter your email.");return;}
  btn.disabled=true;
  google.script.run.withSuccessHandler(function(){
    btn.disabled=false;_pendingEmail=email;
    document.getElementById("resetErr").classList.remove("show");
    showAuthScreen("reset");
  }).withFailureHandler(function(e){btn.disabled=false;showErr("forgotErr","Server error: "+(e&&e.message?e.message:""));}).requestPasswordResetFromWeb(email);
}
function doReset(){
  var otp=document.getElementById("resetOtp").value.trim();
  var p1=document.getElementById("resetPass").value,p2=document.getElementById("resetPass2").value;
  var btn=document.getElementById("resetBtn");
  document.getElementById("resetErr").classList.remove("show");
  if(!otp||!p1){showErr("resetErr","Fill in all fields.");return;}
  if(p1!==p2){showErr("resetErr","Passwords do not match.");return;}
  btn.disabled=true;
  google.script.run.withSuccessHandler(function(r){
    btn.disabled=false;
    if(r&&r.success){showAuthScreen("login");document.getElementById("loginEmail").value=_pendingEmail;showErr("loginErr","Password reset — please log in.");document.getElementById("loginErr").style.background="rgba(77,166,255,.12)";document.getElementById("loginErr").style.color="var(--ac)";}
    else{showErr("resetErr",(r&&r.error)||"Reset failed.");}
  }).withFailureHandler(function(e){btn.disabled=false;showErr("resetErr","Server error: "+(e&&e.message?e.message:""));}).resetPasswordFromWeb(_pendingEmail,otp,p1,p2);
}
function doLogout(){
  google.script.run.withFailureHandler(function(){}).logoutFromWeb(_authToken);
  localStorage.removeItem("sdToken");_authToken=null;
  closeProfile();
  location.reload();
}

/* ══ TOP TAB ROUTER — Watchlist/Portfolio are now tabs, not side panels.
   switchTab() itself is completely untouched; this just decides whether
   to show the WL/PT view or delegate straight through to it. ══ */

/* XIRR — Newton-Raphson implementation for portfolio return calculation */
function xirr(cashflows, dates){
  if(!cashflows||cashflows.length<2)return null;
  var MAX_ITER=100, TOLS=1e-7, guess=0.1;
  function npv(rate){
    var t0=dates[0].getTime();
    return cashflows.reduce(function(s,cf,i){return s+cf/Math.pow(1+rate,(dates[i].getTime()-t0)/31536e6);},0);
  }
  function dnpv(rate){
    var t0=dates[0].getTime();
    return cashflows.reduce(function(s,cf,i){var t=(dates[i].getTime()-t0)/31536e6;return s-t*cf/Math.pow(1+rate,t+1);},0);
  }
  var r=guess;
  for(var i=0;i<MAX_ITER;i++){
    var f=npv(r),df=dnpv(r);
    if(Math.abs(df)<1e-12)break;
    var r2=r-f/df;
    if(Math.abs(r2-r)<TOLS)return r2;
    r=r2;
    if(r<=-1)r=-0.9999;
  }
  return null;
}
function computePortfolioXIRR(rows){
  if(!rows||!rows.length)return null;
  var cfs=[],dates=[];
  rows.forEach(function(r){
    if(!r.avgBuyPrice||!r.qty)return;
    var d=r.buyDate?new Date(r.buyDate):new Date();
    if(isNaN(d))d=new Date();
    var dir=r.tradeType==="SELL"?-1:1;
    cfs.push(-r.avgBuyPrice*r.qty*dir); // outflow at purchase
    dates.push(d);
  });
  var totalCurrent=rows.reduce(function(s,r){
    if(!r.ltp||!r.qty)return s;
    var dir=r.tradeType==="SELL"?-1:1;
    return s+(r.avgBuyPrice*r.qty+(r.ltp-r.avgBuyPrice)*r.qty*dir);
  },0);
  if(!cfs.length||totalCurrent<=0)return null;
  cfs.push(totalCurrent); dates.push(new Date());
  return xirr(cfs,dates);
}

function switchTopTab(tab){
  document.querySelectorAll(".tbtn").forEach(function(b){b.classList.toggle("active",b.dataset.tab===tab);});
  closeDrST();closeDrI();closeDrLT();
  _curCategory=(tab==="watchlist")?"watchlist":(tab==="intraday"||tab==="shortterm"||tab==="longterm")?tab:null;
  var wlC=document.getElementById("cnt-watchlist"),ptC=document.getElementById("cnt-portfolio");
  var pbar=document.querySelector(".pbar"),krow=document.querySelector(".krow");
  if(tab==="watchlist"||tab==="portfolio"){
    ["cnt-home","cnt-intraday","cnt-active","cnt-longterm"].forEach(function(id){var e=document.getElementById(id);if(e)e.style.display="none";});
    if(pbar)pbar.style.display="none";if(krow)krow.style.display="none";
    ["dr","idr","ltdr"].forEach(function(id){var e=document.getElementById(id);if(e)e.style.display="none";}); // v3: the empty collapsed drawers still add grid gaps above the header
    wlC.style.display=tab==="watchlist"?"block":"none";
    ptC.style.display=tab==="portfolio"?"block":"none";
    var sp=document.querySelector("#favg .v");if(sp){sp.textContent="-";sp.className="v";}
    ensureNseMaster(function(){});
    if(tab==="watchlist"){loadWL();clearInterval(window._wlTimer);window._wlTimer=setInterval(function(){if(document.getElementById("cnt-watchlist").style.display!=="none")loadWL(true);},30000);}
    else{loadPT();applyTradingAccountPref_();clearInterval(window._ptTimer);window._ptTimer=setInterval(function(){if(document.getElementById("cnt-portfolio").style.display!=="none"){loadPT(true);refreshLiveSilently_();}},30000);}
  } else {
    wlC.style.display="none";ptC.style.display="none";
    clearInterval(window._wlTimer);clearInterval(window._ptTimer);
    if(pbar)pbar.style.display="";if(krow)krow.style.display="";
    ["dr","idr","ltdr"].forEach(function(id){var e=document.getElementById(id);if(e)e.style.display="";});
    switchTab(tab);
  }
}

/* ══ AUTOCOMPLETE — generic, reused for any symbol-entry field.
   Matches on symbol prefix OR company-name substring, sourced from the
   NSE master sheet (fetched once per session, lazily, on first use). ══ */
var _nseMaster=null,_nseFetching=false,_nseFetchStartedAt=0,_acHi=-1;
function refreshNseMaster_(cb){
  // If a fetch genuinely never resolves (rare, but google.script.run doesn't guarantee either
  // handler always fires — e.g. a hung/very slow server execution), _nseFetching could stay true
  // forever, silently freezing this data while everything else on the page keeps refreshing
  // normally around it (footer timestamp included) — exactly the "only the time updates" symptom.
  // Self-heals: a stuck flag older than 20s is treated as stale and overridden.
  if(_nseFetching){
    if(!_nseFetchStartedAt||(Date.now()-_nseFetchStartedAt)<20000){setTimeout(function(){refreshNseMaster_(cb);},200);return;}
  }
  _nseFetching=true;_nseFetchStartedAt=Date.now();
  google.script.run.withSuccessHandler(function(r){
    _nseFetching=false;
    _nseMaster=(r&&r.rows)||[];
    // Pre-compute uppercase name once per row at load time — doing this
    // inside the filter (on every keystroke) was the actual cause of the
    // "very delayed" search: a full synchronous .toUpperCase() over every
    // NSE-listed company name, on every single character typed.
    _nseMaster.forEach(function(r){r._nu=(r.name||"").toUpperCase();});
    if(cb)cb();
  }).withFailureHandler(function(){_nseFetching=false;if(!_nseMaster)_nseMaster=[];if(cb)cb();}).getNseMasterFromWeb();
}
function ensureNseMaster(cb){
  if(_nseMaster){cb();return;}
  refreshNseMaster_(cb);
}
function placeAcBox(inp){
  var box=document.getElementById("acBox"),r=inp.getBoundingClientRect();
  box.style.left=r.left+"px";box.style.width=r.width+"px";box.style.top=(r.bottom+4)+"px";
}
function attachAutocomplete(inputId, onPick){
  var inp=document.getElementById(inputId),box=document.getElementById("acBox"),acTmr=null;
  function render(list){
    _acHi=-1;
    if(!list.length){box.classList.remove("show");box.innerHTML="";return;}
    placeAcBox(inp);
    box.innerHTML=list.slice(0,12).map(function(r,i){
      return '<div class="acItem" data-i="'+i+'" onmousedown="event.preventDefault();pickAc(\''+inputId+'\','+i+')"><span class="acSym">'+eh(r.symbol)+'</span><span class="acName">'+eh(r.name||"")+'</span></div>';
    }).join("");
    box._list=list;box._activeInput=inputId;
    box.classList.add("show");
  }
  inp.addEventListener("input",function(){
    var q=inp.value.trim().toUpperCase();
    clearTimeout(acTmr);
    if(!q){box.classList.remove("show");return;}
    acTmr=setTimeout(function(){
      ensureNseMaster(function(){
        var list=_nseMaster.filter(function(r){
          return r.symbol.indexOf(q)>=0 || r._nu.indexOf(q)>=0;
        }).sort(function(a,b){
          var aP=a.symbol.indexOf(q)===0?0:1, bP=b.symbol.indexOf(q)===0?0:1;
          return aP-bP;
        });
        render(list);
      });
    },120);
  });
  inp.addEventListener("keydown",function(e){
    if(box._activeInput!==inputId)return;
    var items=box.querySelectorAll(".acItem");
    if(!items.length)return;
    if(e.key==="ArrowDown"){e.preventDefault();_acHi=Math.min(_acHi+1,items.length-1);items.forEach(function(it,i){it.classList.toggle("hi",i===_acHi);});}
    else if(e.key==="ArrowUp"){e.preventDefault();_acHi=Math.max(_acHi-1,0);items.forEach(function(it,i){it.classList.toggle("hi",i===_acHi);});}
    else if(e.key==="Enter" && _acHi>=0){e.preventDefault();pickAc(inputId,_acHi);}
  });
  inp.addEventListener("blur",function(){setTimeout(function(){if(box._activeInput===inputId)box.classList.remove("show");},150);});
  window.addEventListener("resize",function(){if(box._activeInput===inputId&&box.classList.contains("show"))placeAcBox(inp);});
  inp._acPick=onPick;
}
function pickAc(inputId,idx){
  var box=document.getElementById("acBox"),inp=document.getElementById(inputId);
  var item=box._list&&box._list[idx];if(!item)return;
  inp.value=item.symbol;
  box.classList.remove("show");
  if(inp._acPick)inp._acPick(item);
}

/* ══ WATCHLIST — full-width tab, sourced from the NSE master sheet ══ */
var _wlRows=[];
function mergeNseData_(rows){
  if(!rows||!rows.length)return rows;
  if(_nseMaster&&_nseMaster.length){
    var map={};_nseMaster.forEach(function(m){map[m.symbol]=m;});
    rows.forEach(function(r){
      var m=map[r.symbol.toUpperCase()];
      if(m){r.name=m.name;r.sector=m.sector;r.industry=m.industry;r.ltp=m.ltp;r.change=m.change;r.changePct=m.changePct;r.open=m.open;r.high=m.high;r.low=m.low;r.prevClose=m.prevClose;r.volume=m.volume;r.volumeAvg=m.volumeAvg;r.week52High=m.week52High;r.week52Low=m.week52Low;r.pe=m.pe;r.eps=m.eps;r.marketCap=m.marketCap;}
      else{r.name=r.symbol;}
    });
  } else {rows.forEach(function(r){r.name=r.symbol;});}
  return rows;
}
function loadWL(silent){
  clearTimeout(window._wlAutoRef);
  if(!silent)document.getElementById("wlTblBdy").innerHTML='<tr class="emr"><td colspan="18">Loading...</td></tr>';
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(!r||!r.success){if(!silent)document.getElementById("wlTblBdy").innerHTML='<tr class="emr"><td colspan="18">Error: '+(r&&r.error?eh(r.error):"unknown")+'</td></tr>';return;}
    _wlRows=mergeNseData_(r.rows||[]);
    _wlSymSet.clear();_wlRows.forEach(function(row){_wlSymSet.add(row.symbol);});
    renderWL();
    // Always refresh NSE master here (was: only once ever) — that froze LTP at whatever loaded first each session; getNseMaster_ itself is still cached ~45s server-side, so this stays cheap.
    refreshNseMaster_(function(){_wlRows=mergeNseData_(_wlRows);renderWL();});
  }).withFailureHandler(function(e){if(!silent)document.getElementById("wlTblBdy").innerHTML='<tr class="emr"><td colspan="18">Server error: '+(e&&e.message?eh(e.message):"unknown")+'</td></tr>';}).getWatchlistFromWeb(_authToken);
}
function wlItemsFor_(rows){
  var items=[];
  var nseErr=rows.find(function(r){return r.nseError;});
  if(nseErr)items.push({key:"__banner__",rows:['<tr><td colspan="18" style="background:rgba(255,180,60,.12);color:var(--am);font-size:10px;padding:8px">Live price/sector data unavailable right now (NSE sheet: '+eh(nseErr.nseError)+'). Your watchlist itself loaded fine.</td></tr>']});
  rows.forEach(function(r,i){
    var chg=r.changePct;
    var chgStr=chg!=null?(chg>=0?"+":"")+chg.toFixed(2)+"%":"-";
    var chgCls=chg!=null?(chg>=0?"rp":"rn"):"";
    var chgAbsStr=r.change!=null?(r.change>=0?"+":"")+fNZ(r.change):"-";
    var html='<tr><td style="color:var(--mu)">'+(i+1)+'</td>'+symCell(r.symbol,r.name||"",null,"watchlist",r.notes)+
      '<td style="color:var(--dm)">'+eh(r.sector||"-")+'</td><td style="color:var(--dm)">'+eh(r.industry||"-")+'</td>'+
      '<td>'+fNZ(r.ltp)+'</td><td class="'+chgCls+'">'+chgAbsStr+'</td><td class="'+chgCls+'">'+chgStr+'</td>'+
      '<td>'+fNZ(r.open)+'</td><td>'+fNZ(r.high)+'</td><td>'+fNZ(r.low)+'</td><td>'+fNZ(r.prevClose)+'</td>'+
      '<td>'+fCr(r.volume)+'</td><td>'+fCr(r.volumeAvg)+'</td><td>'+fNZ(r.week52High)+'</td><td>'+fNZ(r.week52Low)+'</td>'+
      '<td>'+fNZ(r.pe)+'</td><td>'+fNZ(r.eps)+'</td><td>'+fCr(r.marketCap)+'</td></tr>';
    items.push({key:r.symbol,rows:[html]});
  });
  return items;
}
function renderWL(){
  var b=document.getElementById("wlTblBdy");
  document.getElementById("wlCount").textContent=_wlRows.length;
  if(!_wlRows.length){b.innerHTML='<tr class="emr"><td colspan="18">Your watchlist is empty. Add a symbol above.</td></tr>';updFoot([],function(){return null;});return;}
  // A background refresh (loadWL(true), every 30s) calls this too, along with the NSE-merge
  // callback right after it — if a search filter is active, re-render through it instead of the
  // full list, or the filtered view reverts to showing everything on every refresh regardless of
  // what was typed, which is what was reported as flickering/unstable while filtering.
  var q=(document.getElementById("wlSearch")||{}).value||"";
  if(q.trim())filterWL(); else patchTbody(b,wlItemsFor_(_wlRows));
  updFoot(_wlRows,function(r){return r.changePct!=null?r.changePct/100:null;});
}
function addWLItem(){
  var sym=document.getElementById("wlSym").value.trim();
  var note=document.getElementById("wlNote").value.trim();
  var st=document.getElementById("wlAddStatus");
  if(!sym)return;
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){
      document.getElementById("wlSym").value="";document.getElementById("wlNote").value="";
      loadWL(true);
      if(st){st.textContent="\u2713 Added";st.style.color="var(--gr)";clearTimeout(st._fadeT);st._fadeT=setTimeout(function(){st.textContent="";},2500);}
    }
    else if(r&&r.error){if(st){st.textContent="Error: "+r.error;st.style.color="var(--rd)";}}
  }).withFailureHandler(function(e){if(st){st.textContent="Server error: "+(e&&e.message?e.message:"unknown error");st.style.color="var(--rd)";}}).addToWatchlistFromWeb(_authToken,sym,note);
}
function removeWLItem(rowId){
  google.script.run.withSuccessHandler(function(r){if(r&&r.sessionExpired){forceReauth();return;}loadWL(true);}).withFailureHandler(function(){}).removeFromWatchlistFromWeb(_authToken,rowId);
}

/* ══ PORTFOLIO — full-width tab, LTP sourced from the NSE master sheet ══ */
var _ptRows=[];
var _ptXirrBySymbol={};
var _ptSubtab="holdings";
var _allTradesBySymbol=null;
function loadPT(silent){
  if(!silent){
    document.getElementById("ptTblBdy").innerHTML='<tr class="emr"><td colspan="13">Loading...</td></tr>';
    document.getElementById("posTblBdy").innerHTML='<tr class="emr"><td colspan="12">Loading...</td></tr>';
    _ptExpanded={};
  }
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(!r||!r.success){if(!silent)document.getElementById("ptTblBdy").innerHTML='<tr class="emr"><td colspan="13">Error: '+(r&&r.error?eh(r.error):"unknown")+'</td></tr>';return;}
    // v2: the brokerage-charges settings, margin and the user's full trade list now arrive WITH the portfolio,
    // so net P&L and the "traded today" Positions rows are right on the very first paint.
    if(r.brokerage)_brk=r.brokerage;
    if(r.margin)_marginInfo=r.margin;
    if(r.tradesBySymbol)_allTradesBySymbol=r.tradesBySymbol;
    _netMemo={};
    _ptRows=mergeNseData_(r.rows||[]);
    recomputePtXirr_();
    renderPT();
    renderMarginStrips_();
    if(!r.tradesBySymbol)loadPortfolioXirr_();
    Object.keys(_ptExpanded).forEach(function(pk){
      if(!_ptExpanded[pk])return;
      var k=pk.replace(/^pos\|/,"");
      var sep=k.lastIndexOf("__");
      loadTradesFor(k.slice(0,sep),k.slice(sep+2),true);
    });
    // Always refresh NSE master here (was: only once ever) — see loadWL for why.
    refreshNseMaster_(function(){_ptRows=mergeNseData_(_ptRows);_netMemo={};renderPT();});
  }).withFailureHandler(function(e){if(!silent)document.getElementById("ptTblBdy").innerHTML='<tr class="emr"><td colspan="13">Server error: '+(e&&e.message?eh(e.message):"unknown")+'</td></tr>';}).getPortfolioFromWeb(_authToken);
}
function loadPortfolioXirr_(){
  google.script.run.withSuccessHandler(function(r){
    if(!r||!r.success)return;
    _allTradesBySymbol=r.bySymbol||{};
    _netMemo={};
    recomputePtXirr_();
    renderPT();
    if(_ptSubtab==="tradelog")renderTradeLog();
  }).withFailureHandler(function(){}).getAllPortfolioTradesFromWeb(_authToken);
}
function recomputePtXirr_(){
  if(!_allTradesBySymbol)return;
  _ptXirrBySymbol={};
  _ptRows.forEach(function(posRow){
    var pk=posRow.symbol+"__"+posRow.category;
    var ownCategoryTrades=(_allTradesBySymbol[posRow.symbol]||[]).filter(function(t){return t.category===posRow.category;});
    // XIRR for a currently-held position should reflect just this open cycle in its OWN category,
    // not get blended with an unrelated cycle from a different category holding the same symbol
    // (e.g. a stock traded both intraday and long-term) or an already fully-closed cycle from the
    // same symbol's past — extracted here client-side; the full unfiltered list from the server is
    // what the Trade Log itself uses (split by category there too — see tlGroupsAll_).
    _ptXirrBySymbol[pk]=computeStockXIRR(currentCycleTrades_(ownCategoryTrades),posRow.ltp);
  });
}
function currentCycleTrades_(trades){
  var t=(trades||[]).slice().sort(function(a,b){return new Date(a.tradeDate||0)-new Date(b.tradeDate||0);});
  var netQty=0,cutIdx=0;
  t.forEach(function(tr,i){netQty+=(tr.direction==="SELL"?-1:1)*tr.qty;if(Math.round(netQty*100)===0)cutIdx=i+1;});
  return t.slice(cutIdx);
}
function isTodayLocal_(d){
  if(!d)return false;
  var a=new Date(d),n=new Date();
  return a.getFullYear()===n.getFullYear()&&a.getMonth()===n.getMonth()&&a.getDate()===n.getDate();
}
function splitPtRows_(){
  var holdings=[],positions=[];
  _ptRows.forEach(function(r){
    // Intraday always sits in Positions and never migrates to Holdings — it's meant to be fully
    // exited same-day (see closeIntradayAtMarketEnd_/the 3PM force-sell), so there's no legitimate
    // "next day" state for it to hold. Short/long-term bought or ADDED TO today sit in Positions and
    // move to Holdings once lastTradeDate is no longer today (a plain date comparison evaluated
    // fresh on every load, so it just naturally happens at midnight — no separate "move it over"
    // step needed). This used to check buyDate (the position's very first trade, ever) — so an
    // order-book execution or a manual top-up on a holding you already had never showed here at
    // all, even though the trade itself was recorded correctly: buyDate simply never moves off the
    // original purchase date. lastTradeDate is what actually answers "did anything happen today".
    if(r.category==="intraday"){positions.push(r);return;}
    if(isTodayLocal_(r.lastTradeDate||r.buyDate)){
      // v2: a stock you ALREADY held that traded again today stays in Holdings (it is still a holding) AND gets a
      // Positions row for today's activity. A position opened for the first time today (or one fully sold today)
      // is Positions-only, exactly as before.
      var sl=(r.qty>0)?todaySlice_(r):null;
      if(sl&&sl.hadPriorQty){holdings.push(r);positions.push(sl.row);}
      else positions.push(r);
    } else holdings.push(r);
  });
  return {holdings:holdings,positions:positions};
}
function summarize_(rows){
  // v2: built on rowFigures_ so the header, the rows and the Excel export always agree, and so
  // brokerage & charges are included when enabled (Settings → Brokerage).
  var totInvested=0,totCurrent=0,gain=0,charges=0;
  rows.forEach(function(r){
    // A closed-out row has qty 0, so the live-mark formula would contribute nothing at all —
    // rowFigures_ returns its realized figures explicitly, so the day's banked P&L never vanishes
    // from the footer the moment a position is squared off.
    var f=rowFigures_(r);
    if(f.invested==null)return;
    totInvested+=f.invested;
    totCurrent+=f.curVal!=null?f.curVal:f.invested;
    gain+=f.pnl!=null?f.pnl:0;
    charges+=f.charges||0;
  });
  return {invested:totInvested,current:totCurrent,gain:gain,charges:charges};
}
function tlFooterStats_(){
  var groups=tlGroupsAll_();
  var invested=0,gain=0;
  groups.forEach(function(g){invested+=g.invested||0;gain+=g.pnl||0;});
  return {invested:invested,gain:gain};
}
function updatePtFooter_(){
  var fx;
  if(_tradingMode==="live")fx=liveFooterStats_(); // v2: Live shows its own live holdings/positions only - never the mock portfolio's average
  else{
    var split=splitPtRows_();
    if(_ptSubtab==="positions")fx=summarize_(split.positions);
    else if(_ptSubtab==="tradelog")fx=tlFooterStats_();
    else fx=summarize_(split.holdings);
  }
  var sp=document.querySelector("#favg .v");
  if(!fx||!fx.invested){if(sp){sp.textContent="-";sp.className="v";}return;}
  if(sp){var wr=fx.invested>0?(fx.gain/fx.invested):null;sp.textContent=wr!=null?fR(wr):"-";sp.className="v "+(wr!=null&&wr>=0?"pos":"neg");}
  document.getElementById("ftm").textContent=new Date().toLocaleTimeString("en-IN",{hour12:false});
}
function renderPT(){
  var split=splitPtRows_();
  document.getElementById("ptCount").textContent=split.holdings.length;
  document.getElementById("posCount").textContent=split.positions.length;

  var sumEl=document.getElementById("ptSummary");
  if(!split.holdings.length){document.getElementById("ptTblBdy").innerHTML='<tr class="emr"><td colspan="13">No holdings yet. Add one above.</td></tr>';if(sumEl)sumEl.innerHTML="";}
  else{
    if(filterActive_("ptSearch","ptCatFilter"))filterPT(); else renderPT_rows(split.holdings);
    var hs=summarize_(split.holdings);
    var xirrVal=computePortfolioXIRR(split.holdings);
    var xirrStr=xirrVal!=null?(" &middot; XIRR: <b class=\""+(xirrVal>=0?"rp":"rn")+"\">"+(xirrVal>=0?"+":"")+(xirrVal*100).toFixed(1)+"%</b>"):"";
    // v2: Gain/Loss broken out by timeframe here too (as Positions already did)
    var hCat=catGainLine_(split.holdings.map(function(r){return{category:r.category,gain:rowFigures_(r).pnl||0};}));
    if(sumEl)sumEl.innerHTML='Invested: <b>&#8377;'+hs.invested.toLocaleString("en-IN",{maximumFractionDigits:0})+'</b> &middot; '+
      'Current: <b>&#8377;'+hs.current.toLocaleString("en-IN",{maximumFractionDigits:0})+'</b> &middot; '+
      'Gain/Loss: <b class="'+(hs.gain>=0?"rp":"rn")+'">'+(hs.gain>=0?"+":"")+'&#8377;'+hs.gain.toLocaleString("en-IN",{maximumFractionDigits:0})+'</b>'+chargesHtml_(hs)+xirrStr+
      (hCat?' &middot; '+hCat:'');
  }

  var posSumEl=document.getElementById("posSummary");
  if(!split.positions.length){document.getElementById("posTblBdy").innerHTML='<tr class="emr"><td colspan="12">No intraday positions today.</td></tr>';if(posSumEl)posSumEl.innerHTML="";}
  else{
    if(filterActive_("posSearch","posCatFilter"))filterPositions(); else renderPositions_rows(split.positions);
    var ps=summarize_(split.positions);
    // Positions can mix intraday with today's short/long-term activity — a single blended
    // Gain/Loss doesn't say how each timeframe actually did today, so break it out per category
    // too (only for whichever categories actually have rows right now).
    var byCat={};
    split.positions.forEach(function(r){(byCat[r.category]=byCat[r.category]||[]).push(r);});
    var catLine=["intraday","shortterm","longterm"].map(function(cat){
      if(!byCat[cat]||!byCat[cat].length)return null;
      var cs=summarize_(byCat[cat]);
      var label={intraday:"Intraday",shortterm:"Short-term",longterm:"Long-term"}[cat];
      return label+': <b class="'+(cs.gain>=0?"rp":"rn")+'">'+(cs.gain>=0?"+":"")+'&#8377;'+cs.gain.toLocaleString("en-IN",{maximumFractionDigits:0})+'</b>';
    }).filter(Boolean).join(' &middot; ');
    if(posSumEl)posSumEl.innerHTML='Invested: <b>&#8377;'+ps.invested.toLocaleString("en-IN",{maximumFractionDigits:0})+'</b> &middot; '+
      'Current: <b>&#8377;'+ps.current.toLocaleString("en-IN",{maximumFractionDigits:0})+'</b> &middot; '+
      'Gain/Loss: <b class="'+(ps.gain>=0?"rp":"rn")+'">'+(ps.gain>=0?"+":"")+'&#8377;'+ps.gain.toLocaleString("en-IN",{maximumFractionDigits:0})+'</b>'+chargesHtml_(ps)+
      (catLine?' &middot; '+catLine:'');
  }

  updatePtFooter_();
}
function updPtEst(){
  var qty=Number(document.getElementById("ptQty").value)||0;
  var price=Number(document.getElementById("ptPrice").value)||0;
  var type=document.getElementById("ptType").value;
  var el=document.getElementById("ptEst");
  if(!qty||!price){el.textContent="";return;}
  var amt=(qty*price).toLocaleString("en-IN",{maximumFractionDigits:0});
  el.textContent=(type==="SELL"?"Estimated proceeds from short sale: \u20B9":"Estimated investment: \u20B9")+amt;
}
function addPTItem(){
  var sym=document.getElementById("ptSym").value.trim();
  var qty=document.getElementById("ptQty").value;
  var price=document.getElementById("ptPrice").value;
  var date=document.getElementById("ptDate").value;
  var type=document.getElementById("ptType").value;
  var cat=document.getElementById("ptCat").value;
  if(!sym||!qty||!price)return;
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){document.getElementById("ptSym").value="";document.getElementById("ptQty").value="";document.getElementById("ptPrice").value="";document.getElementById("ptEst").textContent="";loadPT(true);}
    else if(r&&r.error){alert(r.error);}
  }).withFailureHandler(function(e){alert("Could not add trade: "+(e&&e.message?e.message:"unknown error"));}).addPortfolioTradeFromWeb(_authToken,sym,type,qty,price,date,"","Manual",cat);
}
function removePTItem(symbol,category){
  var catLabel={intraday:"Intraday",shortterm:"Short-term",longterm:"Long-term"}[category]||category;
  if(!confirm("Remove this "+catLabel+" position for "+symbol+"? This deletes all "+catLabel+" trades for this symbol (other categories holding "+symbol+" are not affected)."))return;
  google.script.run.withSuccessHandler(function(r){if(r&&r.sessionExpired){forceReauth();return;}loadPT(true);}).withFailureHandler(function(){}).removePortfolioPositionFromWeb(_authToken,symbol,category);
}

/* ══ ORDER BOOK — place orders from any tab via the order-book icon next
   to a symbol; edit/cancel from the Watchlist tab's Order Book section. ══ */
var _obSym="",_obCat="";
var _obAcct="mock";
function openOrderModal(sym,category,ltp){
  _obSym=sym;_obCat=category;
  document.getElementById("obSym").textContent=sym;
  document.getElementById("obCat").textContent={intraday:"Intraday",shortterm:"Short-term",longterm:"Long-term",watchlist:"Watchlist"}[category]||category;
  document.getElementById("obType").value="BUY";
  document.getElementById("obQty").value="";document.getElementById("obPrice").value="";
  document.getElementById("obTarget").value="";document.getElementById("obSL").value="";document.getElementById("obNotes").value="";
  // v4: manual orders from the live panel can go to the Mock account or (real money) to the Live account, per Trading Account
  var av=obAvail_();
  _obAcct=av==="live"?"live":(av==="both"&&_tradingMode==="live"?"live":"mock");
  document.getElementById("obAcctRow").style.display=av==="both"?"":"none";
  document.getElementById("obKind").value="LIMIT";
  applyOrderAcct_();
  document.getElementById("orderM").classList.add("open");
  var badge=document.getElementById("obLtpBadge");
  if(ltp&&Number(ltp)>0){badge.textContent="(LTP: "+fNZ(Number(ltp))+")";document.getElementById("obPrice").placeholder=String(Number(ltp));}
  else{badge.textContent="";document.getElementById("obPrice").placeholder="";}
  document.getElementById("obEstimate").textContent="";
  updOrderEstimate();
}
function setOrderAcct(a){_obAcct=a;applyOrderAcct_();updOrderEstimate();}
function applyOrderAcct_(){
  var live=_obAcct==="live";
  document.getElementById("obAcctMock").style.cssText="border-color:"+(live?"var(--bd)":"var(--ac)")+";color:"+(live?"var(--dm)":"var(--ac)");
  document.getElementById("obAcctLive").style.cssText="border-color:"+(live?"var(--rd)":"var(--bd)")+";color:"+(live?"var(--rd)":"var(--dm)");
  document.getElementById("obMockOnly").style.display=live?"none":"";
  document.getElementById("obKindRow").style.display=live?"":"none";
  document.getElementById("obNote").innerHTML=live?'&#9888; This sends a <b>real order to Zerodha</b> through your relay. It appears in the <b>Live</b> side of your Order Book.':'Orders are tracked in your Order Book. This does not place a live trade with any broker.';
  document.getElementById("obSubmit").textContent=live?"Place LIVE Order":"Place Order";
}
function updOrderEstimate(){
  var q=Number(document.getElementById("obQty").value)||0;
  var p=Number(document.getElementById("obPrice").value)||0;
  var el=document.getElementById("obEstimate");
  if(!el)return;
  if(!(q&&p)){el.textContent="";return;}
  var txt="Estimated value: ₹"+(q*p).toLocaleString("en-IN",{maximumFractionDigits:0});
  // v2: margin privilege - an intraday order only blocks value/leverage of margin (Settings → Brokerage)
  var side=(document.getElementById("obType")||{}).value||"BUY";
  if(_marginInfo&&_marginInfo.margin>0){
    var lev=(_obCat==="intraday")?(_marginInfo.leverage||1):1;
    txt+=" · Margin needed: ₹"+fmtN_(q*p/lev)+(lev>1?" ("+lev+"x leverage)":"")+" · Available: ₹"+fmtN_(_marginInfo.available);
  }
  var ch=calcCharges_(_obCat,side,q,p);
  if(ch>0)txt+=" · Est. charges: ₹"+ch.toFixed(2);
  el.textContent=txt;
}
var _obEditId="";
function closeOrderModal(){document.getElementById("orderM").classList.remove("open");_obEditId="";var t=document.getElementById("obType");if(t)t.disabled=false;}
// v5: a PENDING mock order can be changed (quantity, price, target, stop-loss, notes) any time before it executes
function openEditOrder(id){
  var o=_orderRows.filter(function(x){return x.orderId===id;})[0];
  if(!o||o.status!=="PENDING"){alert("Only pending orders can be changed.");return;}
  openOrderModal(o.symbol,o.category,o.price);
  _obAcct="mock";document.getElementById("obAcctRow").style.display="none";applyOrderAcct_();
  _obEditId=id;
  var t=document.getElementById("obType");t.value=o.orderType==="SELL"?"SELL":"BUY";t.disabled=true;
  document.getElementById("obQty").value=o.qty;document.getElementById("obPrice").value=o.price;
  document.getElementById("obTarget").value=o.target||"";document.getElementById("obSL").value=o.stopLoss||"";document.getElementById("obNotes").value=o.notes||"";
  document.getElementById("obSubmit").textContent="Update Order";
  document.getElementById("obNote").textContent="You can change a pending order any time before it executes.";
  updOrderEstimate();
}
function submitOrder(){
  var type=document.getElementById("obType").value;
  var qty=document.getElementById("obQty").value,price=document.getElementById("obPrice").value;
  if(_obAcct==="live"){
    var kind=document.getElementById("obKind").value;
    if(!qty||Number(qty)<1){alert("Enter a quantity.");return;}
    if(kind==="LIMIT"&&!(Number(price)>0)){alert("A limit order needs a price.");return;}
    if(!confirm("Send a REAL "+type+" order for "+qty+" \u00D7 "+_obSym+" ("+(kind==="LIMIT"?"LIMIT @ \u20B9"+price:"MARKET")+") to Zerodha?"))return;
    google.script.run.withSuccessHandler(function(r){
      if(r&&r.sessionExpired){forceReauth();return;}
      if(r&&r.success){closeOrderModal();if(r.note)alert(r.note);if(_obPanelOpen&&_obMode==="live")loadOrders(true);if(_tradingMode==="live")loadLivePortfolio();}
      else{alert("Could not place live order: "+(r&&r.error?r.error:"unknown error")+(r&&r.needsReconnect?"\nOpen Portfolio \u2192 Live and press Connect Zerodha first.":""));}
    }).withFailureHandler(function(e){alert("Could not place live order: "+(e&&e.message?e.message:"unknown error"));}).kitePlaceOrderFromWeb(_authToken,_obSym,_obCat==="watchlist"?"longterm":_obCat,type,qty,kind==="LIMIT"?price:"",kind);
    return;
  }
  var target=document.getElementById("obTarget").value,sl=document.getElementById("obSL").value;
  var notes=document.getElementById("obNotes").value.trim();
  if(!qty||!price){alert("Enter quantity and price.");return;}
  if(_obEditId){
    google.script.run.withSuccessHandler(function(r){
      if(r&&r.sessionExpired){forceReauth();return;}
      if(r&&r.success){closeOrderModal();loadOrders(true);}
      else{alert("Could not change the order: "+(r&&r.error?r.error:"unknown error"));}
    }).withFailureHandler(function(e){alert("Could not change the order: "+(e&&e.message?e.message:"unknown error"));}).updateOrderFromWeb(_authToken,_obEditId,qty,price,target,sl,"PENDING",notes);
    return;
  }
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){closeOrderModal();if(_obPanelOpen)loadOrders(true);}
    else{alert("Could not place order: "+(r&&r.error?r.error:"unknown error — no response from server"));}
  }).withFailureHandler(function(e){alert("Could not place order: "+(e&&e.message?e.message:"unknown error"));}).placeOrderFromWeb(_authToken,_obSym,_obCat,type,qty,price,target,sl,notes);
}
/* ══ ORDER BOOK — top-edge slide-down panel ══ */
var _orderRows=[],_obRefreshTimer=null,_obPanelOpen=false;
function openOB(){
  if(_obPanelOpen)return;
  _obPanelOpen=true;
  // v4: the Order Book follows the Portfolio's mode when both accounts are on; Live-only users see only live orders
  if(obAvail_()==="both")_obMode=(_tradingMode==="live")?"live":"mock";
  applyObMode_();
  document.getElementById("obPanel").classList.add("open");
  document.getElementById("obOverlay").classList.add("show");
  loadOrders();
  _obRefreshTimer=setInterval(function(){loadOrders(true);},60000);
}
function closeOB(){
  _obPanelOpen=false;
  document.getElementById("obPanel").classList.remove("open");
  document.getElementById("obOverlay").classList.remove("show");
  clearInterval(_obRefreshTimer);_obRefreshTimer=null;
}
function loadOrders(silent){
  var b=document.getElementById("obTblBdy");if(!b)return;
  if(_obMode==="live"){
    if(!silent)b.innerHTML='<tr class="emr"><td colspan="10">Loading live orders...</td></tr>';
    google.script.run.withSuccessHandler(function(r){
      if(r&&r.sessionExpired){forceReauth();return;}
      if(!r||!r.success){b.innerHTML='<tr class="emr"><td colspan="10">'+eh(r&&r.error?r.error:"Could not load live orders.")+'</td></tr>';return;}
      _liveOrderMeta={ok:!!r.brokerOk,err:r.brokerError||"",reconnect:!!r.needsReconnect};
      _liveOrderRows=r.rows||[];renderLiveOrderBook_();
    }).withFailureHandler(function(e){b.innerHTML='<tr class="emr"><td colspan="10">Server error: '+eh(e&&e.message?e.message:"unknown")+'</td></tr>';}).getLiveOrderBookFromWeb(_authToken);
    return;
  }
  if(!silent)b.innerHTML='<tr class="emr"><td colspan="12">Loading...</td></tr>';
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(!r||!r.success){
      var raw=r===undefined?"undefined (server returned nothing)":r===null?"null":JSON.stringify(r);
      b.innerHTML='<tr class="emr"><td colspan="12">Error: '+eh(r&&r.error?r.error:"Server response: "+raw+" \u2014 if this persists, the deployed web app version may be out of date. Re-deploy (Manage deployments \u2192 Edit \u2192 New version) after pasting in the latest Code.gs.")+'</td></tr>';
      return;
    }
    _orderRows=r.rows||[];renderOrders();
  }).withFailureHandler(function(e){b.innerHTML='<tr class="emr"><td colspan="12">Server error: '+eh(e&&e.message?e.message:"unknown")+'</td></tr>';}).getOrdersFromWeb(_authToken);
}
function renderOrders(rows){
  var b=document.getElementById("obTblBdy");if(!b)return;
  if(rows===undefined){
    // Called with no explicit rows — from a background refresh, or after cancel/delete/mark-
    // executed. If any filter is currently set, reapply it instead of showing the full list, same
    // reasoning as renderWL() above.
    var q=(document.getElementById("obSearch")||{}).value||"";
    var st=(document.getElementById("obFilterStatus")||{}).value||"";
    var cat=(document.getElementById("obFilterCat")||{}).value||"";
    if(q.trim()||st||cat){filterOrders();return;}
  }
  var data=rows||_orderRows;
  document.getElementById("obCount").textContent=_orderRows.length;
  if(!data.length){b.innerHTML='<tr class="emr"><td colspan="13">No orders'+(rows&&rows.length<_orderRows.length?' matching filter':'')+'.</td></tr>';return;}
  var catLabel={intraday:"Intraday",shortterm:"Short-term",longterm:"Long-term",watchlist:"Watchlist"};
  var items=data.map(function(o,i){
    var typeBdg=o.orderType==="SELL"?'<span class="cbdg" style="background:rgba(255,90,90,.15);color:var(--rd)">SELL</span>':'<span class="cbdg" style="background:rgba(70,200,140,.15);color:var(--gr)">BUY</span>';
    var statClr=o.status==="CANCELLED"?"var(--dm)":(o.status==="FILLED"||o.status==="EXECUTED")?"var(--gr)":"var(--am)";
    var dt=o.createdAt?new Date(o.createdAt).toLocaleDateString("en-IN")+' '+new Date(o.createdAt).toLocaleTimeString("en-IN",{hour12:false,hour:"2-digit",minute:"2-digit"}):"";
    var html='<tr><td style="color:var(--mu)">'+(i+1)+'</td><td style="font-weight:700">'+eh(o.symbol)+'</td>'+
      '<td style="color:var(--dm)">'+eh(catLabel[o.category]||o.category)+'</td><td>'+typeBdg+'</td>'+
      '<td>'+o.qty+'</td><td>'+fNZ(o.price)+'</td>'+
      '<td>\u20B9'+Math.round((o.qty||0)*(o.price||0)).toLocaleString("en-IN")+'</td>'+
      '<td>'+(o.target?fNZ(o.target):"-")+'</td>'+
      '<td>'+(o.stopLoss?fNZ(o.stopLoss):"-")+'</td>'+
      '<td style="color:'+statClr+';font-weight:600">'+eh(o.status)+'</td>'+
      '<td style="color:var(--dm);font-size:10px">'+dt+'</td>'+
      '<td>'+(noteIco(o.notes)||'<span style="color:var(--dm)">-</span>')+'</td>'+
      '<td style="white-space:nowrap">'+
        (o.status==="PENDING"?('<button class="rmBtn" title="Change this pending order" onclick="openEditOrder(\''+o.orderId+'\')">&#9998;</button>'+'<button class="rmBtn" title="Mark Executed" style="color:var(--gr)" onclick="markExecuted(\''+o.orderId+'\')">&#10003;</button>'+
        '<button class="rmBtn" title="Cancel" onclick="cancelOrder(\''+o.orderId+'\')">&#10005;</button>'):'')+
        '<button class="rmBtn" title="Delete" onclick="deleteOrder(\''+o.orderId+'\')">&#128465;</button>'+
      '</td></tr>';
    return {key:o.orderId,rows:[html]};
  });
  patchTbody(b,items);
}
function filterOrders(){
  if(_obMode==="live"){renderLiveOrderBook_();return;}
  var q=(document.getElementById("obSearch")||{}).value||"";q=q.trim().toUpperCase();
  var st=(document.getElementById("obFilterStatus")||{}).value||"";
  var cat=(document.getElementById("obFilterCat")||{}).value||"";
  var filtered=_orderRows.filter(function(o){
    var mq=!q||(o.symbol||"").indexOf(q)>=0;
    var ms=!st||o.status===st;
    var mc=!cat||o.category===cat;
    return mq&&ms&&mc;
  });
  renderOrders(filtered);
}
function cancelOrder(orderId){
  google.script.run.withSuccessHandler(function(r){if(r&&r.sessionExpired){forceReauth();return;}loadOrders(true);}).withFailureHandler(function(){}).cancelOrderFromWeb(_authToken,orderId);
}
function deleteOrder(orderId){
  if(!confirm("Delete this order permanently?"))return;
  google.script.run.withSuccessHandler(function(r){if(r&&r.sessionExpired){forceReauth();return;}loadOrders(true);}).withFailureHandler(function(){}).deleteOrderFromWeb(_authToken,orderId);
}
function filterWL(){
  var q=(document.getElementById("wlSearch")||{}).value||"";q=q.trim().toUpperCase();
  var b=document.getElementById("wlTblBdy");if(!b)return;
  var rows=q?_wlRows.filter(function(r){return(r.symbol||"").indexOf(q)>=0||(r.name||"").toUpperCase().indexOf(q)>=0;}):_wlRows;
  if(!rows.length){b.innerHTML='<tr class="emr"><td colspan="18">No matching symbols.</td></tr>';return;}
  patchTbody(b,wlItemsFor_(rows));
}
function filterPT(){
  var q=((document.getElementById("ptSearch")||{}).value||"").trim().toUpperCase(),cat=catFilterVal_("ptCatFilter");
  var b=document.getElementById("ptTblBdy");if(!b)return;
  // v2: filters the Holdings list itself (it used to filter ALL rows, so a search here could pull today's intraday
  // positions into the Holdings table) and adds the Intraday / Short-term / Long-term filter.
  var rows=splitPtRows_().holdings.filter(function(r){return rowMatches_(r,q,cat);});
  if(!rows.length){b.innerHTML='<tr class="emr"><td colspan="13">No matching symbols.</td></tr>';return;}
  renderPT_rows(rows);
}
var _ptExpanded={},_ptTradesCache={};
var CAT_LABEL={intraday:"Intraday",shortterm:"Short-term",longterm:"Long-term"};
/* ══════════════════════════════════════════════════════════════════════════
   v2 ADDITIONS (Oct 2026) — brokerage charges, per-trade FIFO P&L, "traded today" Positions slice,
   category filters, header gain/loss by category, margin strips.  New code only; the existing
   functions that call into this are marked "v2:".
   ══════════════════════════════════════════════════════════════════════════ */
var _brk=null,_marginInfo=null,_netMemo={};
function fmtN_(n){return(Number(n)||0).toLocaleString("en-IN",{maximumFractionDigits:0});}
function chgOn_(){return !!(_brk&&_brk.applyToPnl);}
// One executed leg. Same formula as the server's calcCharges_ (Settings → Brokerage supplies the rates).
function calcCharges_(cat,side,qty,price){
  if(!chgOn_())return 0;
  var t=(Number(qty)||0)*(Number(price)||0);if(!(t>0))return 0;
  var c=(cat==="intraday")?_brk.intraday:_brk.delivery;if(!c)return 0;
  var buy=String(side).toUpperCase()==="BUY";
  var brk=t*(c.brokeragePct||0)/100;if(c.brokerageCap>0)brk=Math.min(brk,c.brokerageCap);
  var stt=t*((buy?c.sttBuyPct:c.sttSellPct)||0)/100;
  var txn=t*((_brk.exchange==="BSE"?c.txnBsePct:c.txnNsePct)||0)/100;
  var sebi=t*(c.sebiPerCrore||0)/10000000;
  var ipft=(_brk.exchange==="BSE")?0:t*(c.ipftPerCrore||0)/10000000; // v3: NSE Investor Protection Fund Trust charge
  var stamp=buy?t*(c.stampBuyPct||0)/100:0;
  var gst=(brk+sebi+txn+ipft)*(c.gstPct||0)/100;
  return brk+stt+txn+sebi+ipft+stamp+gst;
}
/* FIFO engine. Realized P&L is booked on the CLOSING trade (a SELL that closes a long, a BUY that
   covers a short), matched against the oldest open lots; an opening trade only shows P&L for the
   quantity that is STILL open (marked at LTP, less its entry charges and the estimated exit charges).
   Every line therefore counts a round trip exactly once, and the lines add up to the stock's total P&L.
   Charges (when enabled) are taken per leg: entry share + exit share on realized, entry share + one
   estimated exit on open quantity. Returns {lines:Map(trade -> line), totals}. */
function fifoAnalyze_(trades,ltp,cat){
  var order=(trades||[]).map(function(t,i){return{t:t,i:i};});
  order.sort(function(a,b){var d=new Date(a.t.tradeDate||0)-new Date(b.t.tradeDate||0);return(d&&!isNaN(d))?d:a.i-b.i;});
  var lines=new Map(),queue=[];
  var tot={realized:0,realizedGross:0,openQty:0,openCost:0,openSide:null,entryOpen:0,exitEst:0,unrealized:null,charges:0,net:0};
  order.forEach(function(o){
    var t=o.t,side=t.direction==="SELL"?"SELL":"BUY",q=Number(t.qty)||0,px=Number(t.price)||0;
    var cTot=calcCharges_(cat,side,q,px),cpu=q>0?cTot/q:0;
    tot.charges+=cTot;
    var L={side:side,qty:q,price:px,charges:cTot,realized:0,realizedGross:0,matchedBase:0,matchedQty:0,openQty:0,openBase:0,unrealized:null,pnl:null,base:0,matches:[],entryCh:0,exitCh:0};
    lines.set(t,L);
    var rem=q;
    while(rem>1e-9&&queue.length&&queue[0].side!==side){
      var h=queue[0],m=Math.min(rem,h.left);
      var gross=(h.side==="BUY")?m*(px-h.price):m*(h.price-px);
      var net=gross-h.cpu*m-cpu*m;
      L.realized+=net;L.realizedGross+=gross;L.matchedBase+=m*h.price;L.matchedQty+=m;
      L.matches.push({qty:m,price:h.price,date:h.t.tradeDate,side:h.side});L.entryCh+=h.cpu*m;L.exitCh+=cpu*m;
      tot.realized+=net;tot.realizedGross+=gross;
      h.left-=m;rem-=m;
      if(h.left<=1e-9)queue.shift();
    }
    if(rem>1e-9)queue.push({t:t,side:side,left:rem,price:px,cpu:cpu});
  });
  queue.forEach(function(h){tot.openQty+=h.left;tot.openCost+=h.left*h.price;tot.openSide=h.side;tot.entryOpen+=h.cpu*h.left;});
  if(tot.openQty>1e-9&&ltp!=null&&ltp>0){
    tot.exitEst=calcCharges_(cat,tot.openSide==="BUY"?"SELL":"BUY",tot.openQty,ltp); // ONE exit order for the whole open quantity
    tot.unrealized=0;
  }
  queue.forEach(function(h){
    var L=lines.get(h.t);L.openQty=h.left;L.openBase=h.left*h.price;
    if(tot.unrealized!==null){
      var dir=h.side==="BUY"?1:-1;
      L.unrealized=h.left*(ltp-h.price)*dir-h.cpu*h.left-tot.exitEst*(h.left/tot.openQty);
      tot.unrealized+=L.unrealized;
    }
  });
  lines.forEach(function(L){
    var hasM=L.matchedQty>0,hasO=L.openQty>0;
    L.base=L.matchedBase+L.openBase;
    if(hasM||hasO){
      if(hasO&&L.unrealized===null&&!hasM)L.pnl=null;
      else L.pnl=(hasM?L.realized:0)+(L.unrealized||0);
    }
  });
  tot.net=tot.realized+(tot.unrealized||0);
  return{lines:lines,totals:tot};
}
function cycleFifo_(r){
  var all=(_allTradesBySymbol&&_allTradesBySymbol[r.symbol])||[];
  var key=r.symbol+"|"+r.category+"|"+r.ltp+"|"+all.length+"|"+(chgOn_()?1:0);
  if(_netMemo[key])return _netMemo[key];
  var own=all.filter(function(t){return t.category===r.category;});
  return(_netMemo[key]=fifoAnalyze_(currentCycleTrades_(own),r.ltp,r.category));
}
// Charges paid on today's legs of a closed-out row (its realized P&L from the server is gross).
function closedCharges_(r){
  if(!chgOn_())return 0;
  if(!_allTradesBySymbol)return calcCharges_(r.category,"BUY",1,r.realizedInvested||0)+calcCharges_(r.category,"SELL",1,r.realizedExitValue||0);
  var c=0;
  (_allTradesBySymbol[r.symbol]||[]).forEach(function(t){if(t.category===r.category&&isTodayLocal_(t.tradeDate))c+=calcCharges_(r.category,t.direction,t.qty,t.price);});
  return c;
}
// Single source of truth for a Holdings / Positions row's figures (screen, header totals, Excel).
function rowFigures_(r){
  var dir=r.tradeType==="SELL"?-1:1;
  if(r.qty===0&&r.realizedPnl!=null){
    var ch0=r._netFinal?0:closedCharges_(r),p0=r.realizedPnl-ch0;
    return{isClosed:true,invested:r.realizedInvested,curVal:r.realizedExitValue,pnl:p0,pct:r.realizedInvested?p0/r.realizedInvested*100:null,charges:ch0};
  }
  if(!(r.qty>0))return{isClosed:false,invested:null,curVal:null,pnl:null,pct:null,charges:0};
  var invested=r.avgBuyPrice*r.qty;
  if(r.ltp==null)return{isClosed:false,invested:invested,curVal:null,pnl:null,pct:null,charges:0};
  var curVal=invested+(r.ltp-r.avgBuyPrice)*r.qty*dir,gross=curVal-invested,pnl=gross,base=invested,ch=0;
  if(r._slice){
    pnl=(r._sliceUnr!=null?r._sliceUnr:gross)+(r._extraPnl||0);base=invested+(r._extraBase||0);ch=r._sliceCharges||0;
  } else if(chgOn_()){
    var fa=_allTradesBySymbol?cycleFifo_(r):null,entry,exit;
    if(fa&&fa.totals.openQty>0){entry=fa.totals.entryOpen;exit=fa.totals.exitEst;}
    else{entry=calcCharges_(r.category,dir>0?"BUY":"SELL",r.qty,r.avgBuyPrice);exit=calcCharges_(r.category,dir>0?"SELL":"BUY",r.qty,r.ltp);}
    ch=entry+exit;pnl=gross-ch;
  }
  return{isClosed:false,invested:invested,curVal:curVal,pnl:pnl,pct:base?pnl/base*100:null,charges:ch};
}
/* "Traded today" slice. A short/long-term stock you already held that has new trades today stays in
   Holdings (unchanged) AND gets a Positions row describing TODAY's activity: the lots bought today that
   are still open, plus the realized result of anything sold today (FIFO, so a sale consumes the oldest
   shares first). A stock opened for the first time today is Positions-only, exactly as before. */
function todaySlice_(r){
  if(!_allTradesBySymbol||r.category==="intraday")return null;
  var own=(_allTradesBySymbol[r.symbol]||[]).filter(function(t){return t.category===r.category;});
  var todays=own.filter(function(t){return isTodayLocal_(t.tradeDate);});
  if(!todays.length)return null;
  var pre=0;own.forEach(function(t){if(!isTodayLocal_(t.tradeDate))pre+=(t.direction==="SELL"?-1:1)*(Number(t.qty)||0);});
  if(!(pre>0))return{hadPriorQty:false};
  var fa=fifoAnalyze_(own,r.ltp,r.category);
  var oq=0,ob=0,unr=0,unrOk=true,rl=0,rg=0,rb=0,ch=0;
  todays.forEach(function(t){
    var L=fa.lines.get(t);if(!L)return;
    ch+=L.charges;
    if(L.openQty>0){oq+=L.openQty;ob+=L.openBase;if(L.unrealized==null)unrOk=false;else unr+=L.unrealized;}
    if(L.matchedQty>0){rl+=L.realized;rg+=L.realizedGross;rb+=L.matchedBase;}
  });
  var row=Object.assign({},r,{_slice:true,_sliceCharges:ch});
  if(oq>1e-9){row.qty=oq;row.avgBuyPrice=ob/oq;row._sliceUnr=unrOk?unr:null;row._extraPnl=rl;row._extraBase=rb;delete row.realizedPnl;}
  else{row.qty=0;row.realizedPnl=rl;row.realizedInvested=rb;row.realizedExitValue=rb+rg;row._netFinal=true;}
  return{hadPriorQty:true,row:row};
}
/* Category filter (+ search) shared by Holdings / Positions / Trade Log. */
function catFilterVal_(id){var el=document.getElementById(id);return el?el.value:"";}
function filterActive_(searchId,catId){return !!(((document.getElementById(searchId)||{}).value||"").trim()||catFilterVal_(catId));}
function rowMatches_(r,q,cat){
  if(cat&&r.category!==cat)return false;
  if(!q)return true;
  return(r.symbol||"").indexOf(q)>=0||(r.name||"").toUpperCase().indexOf(q)>=0;
}
/* Gain/Loss per category for a list of {category, gain}. */
function catGainLine_(entries){
  var sums={};
  entries.forEach(function(e){sums[e.category]=(sums[e.category]||0)+(e.gain||0);});
  return["intraday","shortterm","longterm"].map(function(cat){
    if(!(cat in sums))return null;
    var g=sums[cat];
    return CAT_LABEL[cat]+': <b class="'+(g>=0?"rp":"rn")+'">'+(g>=0?"+":"")+'&#8377;'+Math.round(g).toLocaleString("en-IN")+'</b>';
  }).filter(Boolean).join(" &middot; ");
}
function chargesHtml_(s){
  return(chgOn_()&&s&&s.charges>0)?' &middot; <span title="Brokerage, STT/CTT, exchange, SEBI, stamp duty and GST - paid so far, plus an estimate for exiting open positions. Edit rates in Settings → Brokerage.">Charges: <b>&#8377;'+fmtN_(s.charges)+'</b></span>':"";
}
/* Available margin strip - mock (all three portfolio sub-tabs) and live. */
function renderMarginStrips_(){
  var el=document.getElementById("ptMarginStrip");if(!el)return;
  var m=_marginInfo;
  if(!m){el.innerHTML="";return;}
  if(!(m.margin>0)){el.innerHTML='Margin not set &middot; <a style="color:var(--ac);cursor:pointer" onclick="openProfile()">set it in Settings</a>';return;}
  var lv=m.leverage>1?' &middot; <b>'+m.leverage+'x</b> intraday':"";
  el.innerHTML='Margin <b>&#8377;'+fmtN_(m.margin)+'</b> &middot; Used <b>&#8377;'+fmtN_(m.used)+'</b>'+
    (m.blocked>0?' <span>(incl. &#8377;'+fmtN_(m.blocked)+' held for pending intraday orders)</span>':"")+
    ' &middot; Available <b class="'+(m.available>0?"rp":"rn")+'">&#8377;'+fmtN_(m.available)+'</b>'+lv;
}
/* Live footer (Avg Return) - computed from the LIVE sub-tab being shown, never from the mock portfolio. */
var _livePosRows=[],_liveHldRows=[];
function liveFooterStats_(){
  var rows=_liveSubtab==="positions"?_livePosRows:(_liveSubtab==="holdings"?_liveHldRows:[]);
  var inv=0,gain=0;
  rows.forEach(function(r){
    var q=Math.abs(Number(r.quantity)||0),avg=Number(r.average_price)||0;
    if(!q||!avg)return;
    inv+=q*avg;gain+=Number(r.pnl)||0;
  });
  return{invested:inv,gain:gain};
}

function inr_(n){return(n>=0?"+":"-")+"\u20B9"+Math.round(Math.abs(n)).toLocaleString("en-IN");}
// Hover text for a trade line's P&L cell.
function lineTip_(L){
  var p=[];
  if(L.matchedQty>0){
    if(chgOn_())p.push("Realized on "+L.matchedQty+" sh: gross "+inr_(L.realizedGross)+" \u2212 charges \u20B9"+(L.entryCh+L.exitCh).toFixed(2)+" (entry leg \u20B9"+L.entryCh.toFixed(2)+" + exit leg \u20B9"+L.exitCh.toFixed(2)+") = "+inr_(L.realized)+" net");
    else p.push("Realized on "+L.matchedQty+" sh: "+inr_(L.realized));
  }
  if(L.openQty>0&&L.unrealized!=null)p.push("Still open "+L.openQty+" sh: "+inr_(L.unrealized)+" unrealized"+(chgOn_()?" after entry charges and an estimate for exiting":""));
  if(chgOn_()&&L.charges&&!(L.matchedQty>0))p.push("This order's charges: \u20B9"+L.charges.toFixed(2));
  return p.join(" \u00B7 ");
}
// XIRR of one trade line: a closing trade is measured against the lots it closed, an opening trade against LTP.
function lineXirr_(L,t,ltp){
  var cfs=[],ds=[];
  if(L.matches&&L.matches.length){
    L.matches.forEach(function(m){
      var d=m.side==="BUY"?1:-1;
      cfs.push(-d*m.qty*m.price);ds.push(new Date(m.date||t.tradeDate));
      cfs.push(d*m.qty*L.price);ds.push(new Date(t.tradeDate));
    });
  } else if(L.openQty>0&&ltp){
    var d2=L.side==="BUY"?1:-1;
    cfs.push(-d2*L.openQty*L.price);ds.push(new Date(t.tradeDate));
    cfs.push(d2*L.openQty*ltp);ds.push(new Date());
  } else return null;
  try{return xirr(cfs,ds);}catch(e){return null;}
}
// Tooltip for a Holdings/Positions P&L cell.
function chgTitle_(fg){
  if(!chgOn_()||!fg||!(fg.charges>0))return"";
  return' title="After brokerage &amp; charges of \u20B9'+Math.round(fg.charges).toLocaleString("en-IN")+(fg.isClosed?"":" (incl. an estimate for exiting)")+'"';
}
// Tooltip for a Trade Log group's P&L cell.
function grpTitle_(g){
  if(g.realized==null)return"";
  var p=["Realized "+inr_(g.realized)];
  if(g.unrealized!=null)p.push("Unrealized "+inr_(g.unrealized));
  if(chgOn_()&&g.charges>0)p.push("after charges of \u20B9"+Math.round(g.charges).toLocaleString("en-IN"));
  return' title="'+eh(p.join(" \u00B7 "))+'"';
}

function expIco_(r,ns){
  var pk=r.symbol+"__"+r.category,ek=(ns==="pos"?"pos|":"")+pk;
  return '<span class="ptExp" data-ek="'+eh(ek)+'" onclick="event.stopPropagation();toggleTradeRow(\''+eh(r.symbol)+'\',\''+eh(r.category)+'\''+(ns==="pos"?',\'pos\'':'')+')" style="cursor:pointer;display:inline-block;width:14px;transform:'+(_ptExpanded[ek]?"rotate(90deg)":"rotate(0deg)")+';transition:transform .15s">&#9656;</span>';
}
function renderPT_rows(data){
  var b=document.getElementById("ptTblBdy");if(!b)return;
  var items=[];
  var nseErr=data.find(function(r){return r.nseError;});
  if(nseErr)items.push({key:"__banner__",rows:['<tr><td colspan="13" style="background:rgba(255,180,60,.12);color:var(--am);font-size:10px;padding:8px">Live price data unavailable ('+eh(nseErr.nseError)+'). Holdings loaded fine.</td></tr>']});
  (data||[]).forEach(function(r,i){
    var pk=r.symbol+"__"+r.category;
    var fg=rowFigures_(r); // v2: net of brokerage & charges when enabled
    var invested=fg.invested,pnl=fg.pnl,pnlPct=fg.pct,curVal=fg.curVal;
    var buyDt=r.buyDate?(new Date(r.buyDate)).toLocaleDateString("en-IN"):"-";
    var pnlStr=pnl!=null?(pnl>=0?"+":"")+pnl.toLocaleString("en-IN",{maximumFractionDigits:0})+(pnlPct!=null?" ("+(pnlPct>=0?"+":"")+pnlPct.toFixed(1)+"%)":""):"-";
    var pnlCls=pnl!=null?(pnl>=0?"rp":"rn"):"";
    var typeBdg=categoryBdg_(r.category);
    var xirrVal=_ptXirrBySymbol[pk];
    var xirrStr=xirrVal!=null?((xirrVal>=0?"+":"")+(xirrVal*100).toFixed(1)+"%"):"-";
    var xirrCls=xirrVal!=null?(xirrVal>=0?"rp":"rn"):"";
    var expIco=expIco_(r,"pt");
    var mainRow='<tr onclick="toggleTradeRow(\''+eh(r.symbol)+'\',\''+eh(r.category)+'\')" style="cursor:pointer"><td style="color:var(--mu)">'+expIco+(i+1)+'</td>'+symCell(r.symbol,r.name||"",null,null,r.notes)+'<td>'+typeBdg+'</td>'+
      '<td>'+r.qty+'</td><td>'+fNZ(r.avgBuyPrice)+'</td><td style="color:var(--dm);font-size:10px">'+buyDt+'</td><td>'+fNZ(r.ltp)+'</td>'+
      '<td>'+(invested!=null?invested.toLocaleString("en-IN",{maximumFractionDigits:0}):"-")+'</td>'+
      '<td>'+(curVal!=null?curVal.toLocaleString("en-IN",{maximumFractionDigits:0}):"-")+'</td>'+
      '<td class="'+pnlCls+'"'+chgTitle_(fg)+'>'+pnlStr+'</td>'+
      '<td class="'+xirrCls+'">'+xirrStr+'</td>'+
      '<td style="color:var(--dm);font-size:10px">'+eh(r.action||"-")+'</td>'+
      '<td style="white-space:nowrap"><button class="rmBtn" title="Remove position" onclick="event.stopPropagation();removePTItem(\''+eh(r.symbol)+'\',\''+eh(r.category)+'\')">&#10005;</button></td></tr>';
    // Render from cache instantly if we have it — avoids a "Loading trades..."
    // flash every time this row gets rebuilt (including on silent background refresh).
    var cached=_ptTradesCache[pk];
    var innerHtml=cached?buildTradesHtml(r.symbol,cached,r.category):"Loading trades...";
    var expRow='<tr id="ptExp_'+eh(pk)+'" style="display:'+(_ptExpanded[pk]?"table-row":"none")+'"><td colspan="13" style="padding:0;background:var(--sf2)"><div id="ptTrades_'+eh(pk)+'" style="padding:10px 16px;font-size:11px;color:var(--dm)">'+innerHtml+'</div></td></tr>';
    items.push({key:pk,rows:[mainRow,expRow]});
  });
  patchTbody(b,items);
}
function renderPositions_rows(data){
  var b=document.getElementById("posTblBdy");if(!b)return;
  var items=(data||[]).map(function(r,i){
    var pk=r.symbol+"__"+r.category;
    var fg=rowFigures_(r); // v2: net of brokerage & charges when enabled; "traded today" rows carry today's slice
    var isClosed=fg.isClosed,invested=fg.invested,pnl=fg.pnl,pnlPct=fg.pct,curVal=fg.curVal;
    var buyDt=(r.lastTradeDate||r.buyDate)?(new Date(r.lastTradeDate||r.buyDate)).toLocaleDateString("en-IN"):"-";
    // A closed-out row shows its actual realized P&L for the day (settled figures, not a mark to a
    // live price that no longer applies to it) — flagged with a "closed" tag so it's still obvious
    // at a glance that the position is no longer running.
    var pnlStr=pnl!=null?((pnl>=0?"+":"")+pnl.toLocaleString("en-IN",{maximumFractionDigits:0})+(pnlPct!=null?" ("+(pnlPct>=0?"+":"")+pnlPct.toFixed(1)+"%)":"")+(isClosed?' <span class="cbdg" style="background:rgba(120,140,170,.15);color:var(--dm);margin-left:4px">CLOSED</span>':"")):"-";
    var pnlCls=pnl!=null?(pnl>=0?"rp":"rn"):"";
    var typeBdg=categoryBdg_(r.category)+(r._slice?'<span class="cbdg" title="Today\'s activity on a stock you already hold - the full holding stays in Holdings" style="background:rgba(77,166,255,.12);color:var(--bl)">TODAY</span>':"");
    var expIco=expIco_(r,"pos");
    var ek="pos|"+pk;
    var mainRow='<tr onclick="toggleTradeRow(\''+eh(r.symbol)+'\',\''+eh(r.category)+'\',\'pos\')" style="cursor:pointer"><td style="color:var(--mu)">'+expIco+(i+1)+'</td>'+symCell(r.symbol,r.name||"",null,"intraday",r.notes)+'<td>'+typeBdg+'</td>'+
      '<td>'+r.qty+'</td><td>'+fNZ(r.avgBuyPrice)+'</td><td style="color:var(--dm);font-size:10px">'+buyDt+'</td><td>'+fNZ(r.ltp)+'</td>'+
      '<td>'+(invested!=null?invested.toLocaleString("en-IN",{maximumFractionDigits:0}):"-")+'</td>'+
      '<td>'+(curVal!=null?curVal.toLocaleString("en-IN",{maximumFractionDigits:0}):"-")+'</td>'+
      '<td class="'+pnlCls+'"'+chgTitle_(fg)+'>'+pnlStr+'</td>'+
      '<td style="color:var(--dm);font-size:10px">'+eh(r.action||"-")+'</td>'+
      '<td style="white-space:nowrap"><button class="rmBtn" title="Remove position" onclick="event.stopPropagation();removePTItem(\''+eh(r.symbol)+'\',\''+eh(r.category)+'\')">&#10005;</button></td></tr>';
    var cached=_ptTradesCache[pk];
    var innerHtml=cached?buildTradesHtml(r.symbol,cached,r.category,{todayOnly:!!r._slice}):"Loading trades...";
    var expRow='<tr id="posExp_'+eh(pk)+'" style="display:'+(_ptExpanded[ek]?"table-row":"none")+'"><td colspan="12" style="padding:0;background:var(--sf2)"><div id="posTrades_'+eh(pk)+'" style="padding:10px 16px;font-size:11px;color:var(--dm)">'+innerHtml+'</div></td></tr>';
    return {key:pk,rows:[mainRow,expRow]};
  });
  patchTbody(b,items);
}
function filterPositions(){
  var q=((document.getElementById("posSearch")||{}).value||"").trim().toUpperCase(),cat=catFilterVal_("posCatFilter");
  var rows=splitPtRows_().positions.filter(function(r){return rowMatches_(r,q,cat);});
  var b=document.getElementById("posTblBdy");if(!b)return;
  if(!rows.length){b.innerHTML='<tr class="emr"><td colspan="12">No matching positions.</td></tr>';return;}
  renderPositions_rows(rows);
}
// Requested: an Excel download for portfolio reports. Built entirely client-side from data this tab
// has already loaded (no extra server round-trip) using SheetJS — one workbook, three sheets
// (Holdings / Positions / Trade Log), mirroring what's on screen in each of those sub-tabs.
function exportPortfolioExcel(){
  if(_tradingMode==="live"){exportLiveExcel_();return;} // v3: the same download icon exports the Live tables when Live is showing
  if(typeof XLSX==="undefined"){alert("The download library didn't load — check your connection and try again.");return;}
  if(_allTradesBySymbol===null){alert("Still loading your trade history — try again in a moment.");loadPortfolioXirr_();return;}
  function positionSheet(rows,withXirr){
    var header=["Symbol","Category","Type","Qty","Avg Price","Buy Date","LTP","Invested","Current Val.","P&L","Return %"].concat(withXirr?["XIRR %"]:[]).concat(["Charges"]);
    var aoa=[header];
    rows.forEach(function(r){
      var fg=rowFigures_(r),invested=fg.invested,curVal=fg.curVal,pnl=fg.pnl,retPct=fg.pct; // v2: same net figures as the screen
      var row=[r.symbol,CAT_LABEL[r.category]||r.category,r.tradeType,r.qty,r.avgBuyPrice,r.buyDate?new Date(r.buyDate).toLocaleDateString("en-IN"):"",r.ltp!=null?r.ltp:"",invested!=null?Math.round(invested):"",curVal!=null?Math.round(curVal):"",pnl!=null?Math.round(pnl):"",retPct!=null?Math.round(retPct*10)/10:""];
      if(withXirr){var x=_ptXirrBySymbol[r.symbol+"__"+r.category];row.push(x!=null?Math.round(x*1000)/10:"");}
      row.push(fg.charges?Math.round(fg.charges):"");
      aoa.push(row);
    });
    return XLSX.utils.aoa_to_sheet(aoa);
  }
  var split=splitPtRows_();
  var wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,positionSheet(split.holdings,true),"Holdings");
  XLSX.utils.book_append_sheet(wb,positionSheet(split.positions,false),"Positions");
  var tlAoa=[["Symbol","Status","Category","Net Qty","Invested","Current/Exit Val.","P&L","Return %","XIRR %"]];
  tlGroupsAll_().forEach(function(g){
    tlAoa.push([g.symbol,g.isOpen?"Running":(g.pnl!=null&&g.pnl>=0?"Success":"Failed"),CAT_LABEL[g.category]||g.category,g.netQty,
      g.invested!=null?Math.round(g.invested):"",g.currentVal!=null?Math.round(g.currentVal):"",g.pnl!=null?Math.round(g.pnl):"",
      g.retPct!=null?Math.round(g.retPct*10)/10:"",g.xirr!=null?Math.round(g.xirr*1000)/10:""]);
  });
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(tlAoa),"Trade Log");
  XLSX.writeFile(wb,"portfolio-report-"+new Date().toISOString().slice(0,10)+".xlsx");
}
function switchPortfolioSubtab(sub){
  _ptSubtab=sub;
  setTimeout(syncStickyOffsets_,0);
  ["holdings","positions","tradelog"].forEach(function(s){
    document.getElementById("ptSub-"+s).style.display=(s===sub)?"":"none";
    document.getElementById("ptSubBtn-"+s).className="lptab"+(s===sub?" ptsub-on":"");
  });
  // v3: the search + timeframe filter of the active sub-tab sit in the tab strip
  document.querySelectorAll("#ptSubTabsBar .ptCtl").forEach(function(el){el.style.display=(el.getAttribute("data-for")===sub)?"":"none";});
  if(sub==="tradelog")renderTradeLog();
  updatePtFooter_();
}
/* ══════════════════════════════════════════════════════════════════════════
   LIVE TRADING (frontend) — window-swap between the existing mock Portfolio
   (untouched, exactly as it always worked) and a real Zerodha/Kite Connect
   view. See ARCHITECTURE_LIVE_TRADING.md for the full design.
   ══════════════════════════════════════════════════════════════════════════ */
var _tradingMode="mock";
var _liveOrdLtp=null; // set when a symbol is picked in the live order form's autocomplete — used for the price watermark and a sanity check before a REAL order goes out
var _tradingAccountPref="mock";
var _liveSubtab="holdings";
var MOCK_PT_IDS=["ptMarginStrip","ptSubTabsBar","ptSub-holdings","ptSub-positions","ptSub-tradelog"];
// Requested: which trading account(s) are available is a standing preference (Trading
// Preferences in Settings), not a control living inside the Portfolio tab itself. 'both' is the
// only case that needs the swap UI at all — 'mock' or 'live' alone just shows that one directly,
// with no mode bar cluttering the tab for users who only ever want one.
function applyTradingAccountPref_(){
  var bar=document.getElementById("ptModeBar");
  if(_tradingAccountPref==="both"){
    bar.style.display="";
    setTradingMode(_tradingMode);
  } else {
    bar.style.display="none";
    setTradingMode(_tradingAccountPref);
  }
}
var _ptAddOpen=false,_liveOrdOpen=false;
function toggleAddHolding(){
  if(_tradingMode==="live"){
    _liveOrdOpen=!_liveOrdOpen;
    document.getElementById("liveOrderPan").style.display=_liveOrdOpen?"":"none";
    if(_liveOrdOpen)setTimeout(function(){var e=document.getElementById("liveOrdSymbol");if(e)e.focus();},50);
  } else {
    _ptAddOpen=!_ptAddOpen;
    document.getElementById("ptAddHoldingPan").style.display=_ptAddOpen?"":"none";
    if(_ptAddOpen)setTimeout(function(){var e=document.getElementById("ptSym");if(e)e.focus();},50);
  }
}
function setTradingMode(mode){
  _tradingMode=mode;
  var isLive=mode==="live";
  MOCK_PT_IDS.forEach(function(id){var el=document.getElementById(id);if(el)el.style.display=isLive?"none":"";});
  // v3: the add-holding form and the live order form keep their own open/closed state (collapsed by default)
  var ap=document.getElementById("ptAddHoldingPan");if(ap)ap.style.display=(!isLive&&_ptAddOpen)?"":"none";
  var op=document.getElementById("liveOrderPan");if(op)op.style.display=(isLive&&_liveOrdOpen)?"":"none";
  var lt=document.getElementById("liveToolbar");if(lt)lt.style.display=isLive?"flex":"none";
  var ab=document.getElementById("ptAddToggleBtn");if(ab){ab.textContent=isLive?"+ Order":"+ Add";ab.title=isLive?"Place a live order":"Add a holding";}
  // The line above shows all three Holdings/Positions/Trade-Log panels at once when returning to
  // Mock — switchPortfolioSubtab is what enforces only one being visible at a time, so it has to
  // run again here, or all three stay stacked visible together (exactly the bug reported).
  if(!isLive)switchPortfolioSubtab(_ptSubtab);
  document.getElementById("liveTradingPanel").style.display=isLive?"":"none";
  document.getElementById("ptModeBanner").style.display=isLive?"":"none";
  document.getElementById("ptModeMockBtn").style.cssText="border-color:"+(isLive?"var(--bd)":"var(--ac)")+";color:"+(isLive?"var(--dm)":"var(--ac)");
  document.getElementById("ptModeLiveBtn").style.cssText="border-color:"+(isLive?"var(--rd)":"var(--bd)")+";color:"+(isLive?"var(--rd)":"var(--dm)");
  if(isLive){switchLiveSubtab(_liveSubtab);checkLiveConnection();}
  updatePtFooter_();renderMarginStrips_(); // v2
}
function checkLiveConnection(){
  var statusEl=document.getElementById("liveConnStatus");
  statusEl.textContent="Checking\u2026";
  var setupBtn=document.getElementById("liveSetupBtn");if(setupBtn)setupBtn.style.display="none";
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.notAllowed){statusEl.textContent=r.error;statusEl.style.color="var(--am)";document.getElementById("liveConnectBtn").style.display="none";if(setupBtn&&r.needsBrokerSetup)setupBtn.style.display="";return;}
    if(r&&r.needsReconnect){
      // kiteAuthHeaders_ found no valid token at all — this is the only case that's genuinely
      // "not connected". Anything else below means a token exists but something else went wrong,
      // which was previously being shown as "Not connected" too, hiding the real reason.
      _liveConnected=false;statusEl.textContent="Not connected";statusEl.style.color="var(--dm)";
      document.getElementById("liveConnectBtn").style.display="";
      tryAutoConnectOnce_(); // v2: if automatic login is switched on, do it now rather than waiting for a manual click
      return;
    }
    document.getElementById("liveConnectBtn").style.display="none"; // a token exists either way past this point
    if(r&&r.success){
      _liveConnected=true;statusEl.textContent="Connected";statusEl.style.color="var(--gr)";
      loadLivePortfolio();loadLiveMargins();loadLiveOrders();
    } else {
      statusEl.textContent="Connected, but: "+(r&&r.error?r.error:"unknown error");statusEl.style.color="var(--am)";
    }
  }).withFailureHandler(function(e){statusEl.textContent="Could not reach the server: "+(e&&e.message?e.message:"unknown");statusEl.style.color="var(--rd)";}).getLiveMarginsFromWeb(_authToken);
}
var _autoConnTried=false,_liveConnected=false;
// v5: Live refreshes every 30 s like the Mock portfolio (positions, holdings, the _LivePortfolios Action mirror, trades and margins)
function refreshLiveSilently_(){if(_tradingMode==='live'&&_liveConnected){loadLivePortfolio();loadLiveMargins();}}
function tryAutoConnectOnce_(){
  if(_autoConnTried)return;_autoConnTried=true;
  google.script.run.withSuccessHandler(function(s){
    if(!s||!s.success||!s.autoConnect||!s.hasPassword||!s.hasTotp)return;
    var statusEl=document.getElementById("liveConnStatus");
    statusEl.textContent="Connecting automatically\u2026";statusEl.style.color="var(--dm)";
    google.script.run.withSuccessHandler(function(r){
      if(r&&r.success){checkLiveConnection();}
      else{statusEl.textContent="Auto-connect failed: "+((r&&r.error)||"unknown")+" \u2014 use Connect Zerodha.";statusEl.style.color="var(--am)";}
    }).withFailureHandler(function(){}).brokerAutoConnectNowFromWeb(_authToken);
  }).withFailureHandler(function(){}).getBrokerSettingsFromWeb(_authToken);
}
function connectBroker(){
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(!r||!r.success){alert(r&&r.error?r.error:"Could not start broker connection.");return;}
    // Opens Kite's own login page. It redirects back to this web app's URL with a request_token,
    // which doGet() exchanges for an access token server-side — nothing to paste back manually.
    var popup=window.open(r.url,"_blank");
    // The OAuth flow completes in that popup, not here — this tab has no way to know that
    // happened on its own (the "Broker connected" page auto-closes itself after 2.5s), so poll
    // until the popup closes, then re-check connection status automatically instead of leaving it
    // showing "Not connected" until a full page reload.
    if(popup){
      var pollTimer=setInterval(function(){
        if(popup.closed){clearInterval(pollTimer);checkLiveConnection();}
      },1000);
      setTimeout(function(){clearInterval(pollTimer);},120000); // stop polling after 2 min regardless, in case the popup never closes
    }
  }).withFailureHandler(function(e){alert("Error: "+(e&&e.message?e.message:"unknown"));}).getKiteLoginUrlFromWeb(_authToken);
}
function loadLiveMargins(){
  google.script.run.withSuccessHandler(function(r){
    if(!r||!r.success)return;
    var d=r.data||{};var avail=d.available||{},used=d.utilised||{};
    // v3: the margin figures live in the toolbar strip (the old separate connection card is gone)
    var strip=document.getElementById("liveMarginStrip");
    if(strip){
      var net=(d.net!=null)?d.net:(avail.live_balance!=null?avail.live_balance:avail.cash);
      strip.innerHTML='Available <b class="rp">&#8377;'+fmtN_(net)+'</b> &middot; Used <b>&#8377;'+fmtN_(used.debits||0)+'</b> &middot; Cash <b>&#8377;'+fmtN_(avail.cash||0)+'</b>';
    }
  }).withFailureHandler(function(){}).getLiveMarginsFromWeb(_authToken);
}
function loadLivePortfolio(){
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(!r||!r.success){
      _livePosRows=[];_liveHldRows=[];_livePosAll=[];_liveHldAll=[];_liveMirror=[];updatePtFooter_();
      var m=eh(r&&r.error?r.error:"Could not load your live portfolio.");
      document.getElementById("livePosTblBdy").innerHTML='<tr class="emr"><td colspan="8">'+m+'</td></tr>';
      document.getElementById("liveHldTblBdy").innerHTML='<tr class="emr"><td colspan="7">'+m+'</td></tr>';
      return;
    }
    _liveMirror=r.mirror||[];
    renderLivePositions((r.positions&&r.positions.net)||[]);
    renderLiveHoldings(r.holdings||[]);
    loadLiveTrades();
  }).withFailureHandler(function(){}).getLivePortfolioFromWeb(_authToken);
}
var _livePosAll=[],_liveHldAll=[],_liveOrdAll=[];
function liveQ_(id){return((document.getElementById(id)||{}).value||"").trim().toUpperCase();}
var _liveMirror=[],_liveTradesBySymbol=null,_liveTlExpanded={},_tlReadOnly=false,_liveOrderRows=[],_liveOrderMeta={ok:true,err:"",reconnect:false};
function liveMirrorMatch_(sym,isIntra){
  sym=String(sym).toUpperCase();
  for(var i=0;i<_liveMirror.length;i++){var m=_liveMirror[i];if(m.symbol===sym&&((m.category==="intraday")===isIntra))return m;}
  return null;
}
// The Action column is computed by YOUR formulas on the mirrored row in _LivePortfolios (same as _Portfolios for Mock)
function liveActionCell_(sym,isIntra){
  var m=liveMirrorMatch_(sym,isIntra),a=m&&m.action?String(m.action).trim():"";
  return a?'<td style="color:var(--dm);font-size:10px">'+eh(a)+'</td>':'<td style="color:var(--dm);font-size:10px">-</td>';
}
function renderLivePositions(rows){
  var open=(rows||[]).filter(function(r){return r.quantity!==0;});
  _livePosAll=open;_livePosRows=open;updatePtFooter_(); // v2: Live footer comes from live rows only
  paintLivePositions_();
}
var _liveRowExpanded={};
function liveTradesFor_(sym,isIntra){
  var all=(_liveTradesBySymbol&&_liveTradesBySymbol[String(sym).toUpperCase()])||[];
  return all.filter(function(t){return(t.category==="intraday")===isIntra;});
}
function liveExpandRow_(key,sym,isIntra,colspan){
  var trades=liveTradesFor_(sym,isIntra),m=liveMirrorMatch_(sym,isIntra),cat=m?m.category:(isIntra?"intraday":"longterm");
  var inner=trades.length?withLiveCtx_(function(){return buildTradesHtml(String(sym).toUpperCase(),trades,cat);}):"No trades recorded for this stock yet. Zerodha only provides today's trades, so this history starts from the day the dashboard first saw them.";
  return '<tr id="lx_'+eh(key)+'" style="display:'+(_liveRowExpanded[key]?"table-row":"none")+'"><td colspan="'+colspan+'" style="padding:0;background:var(--sf2)"><div style="padding:10px 16px;font-size:11px;color:var(--dm)">'+inner+'</div></td></tr>';
}
function liveArrow_(key){return '<span class="ptExp" style="display:inline-block;width:14px;transform:'+(_liveRowExpanded[key]?"rotate(90deg)":"rotate(0deg)")+';transition:transform .15s">&#9656;</span>';}
function toggleLiveRow(key){_liveRowExpanded[key]=!_liveRowExpanded[key];paintLiveSub_();}
function paintLivePositions_(){
  var b=document.getElementById("livePosTblBdy"),all=_livePosAll;
  document.getElementById("livePosCount").textContent=all.length;
  var q=liveQ_("livePosSearch"),pf=liveQ_("livePosFilter");
  var shown=all.filter(function(r){return(!q||String(r.tradingsymbol).toUpperCase().indexOf(q)>=0)&&(!pf||String(r.product).toUpperCase()===pf);});
  var inv=0,cur=0,pnl=0,byP={};
  all.forEach(function(r){var qn=Math.abs(Number(r.quantity)||0),p=Number(r.pnl)||0;inv+=qn*(Number(r.average_price)||0);cur+=qn*(Number(r.last_price)||0);pnl+=p;var k=String(r.product||"").toUpperCase();byP[k]=(byP[k]||0)+p;});
  var sum=document.getElementById("livePosSummary");
  var pl=Object.keys(byP).map(function(k){return({MIS:"Intraday",CNC:"Delivery",NRML:"Normal"}[k]||k)+': <b class="'+(byP[k]>=0?"rp":"rn")+'">'+inr_(byP[k])+'</b>';}).join(" &middot; ");
  sum.innerHTML=all.length?('Invested: <b>&#8377;'+fmtN_(inv)+'</b> &middot; Current: <b>&#8377;'+fmtN_(cur)+'</b> &middot; P&amp;L: <b class="'+(pnl>=0?"rp":"rn")+'">'+inr_(pnl)+'</b>'+(pl?' &middot; '+pl:"")):"";
  if(!all.length){b.innerHTML='<tr class="emr"><td colspan="8">No open positions.</td></tr>';return;}
  if(!shown.length){b.innerHTML='<tr class="emr"><td colspan="8">No matching positions.</td></tr>';return;}
  b.innerHTML=shown.map(function(r,i){
    var pnlCls=r.pnl>=0?"rp":"rn",prod=String(r.product||"").toUpperCase(),intra=prod!=="CNC",key="p|"+String(r.tradingsymbol).toUpperCase()+"|"+prod;
    return '<tr onclick="toggleLiveRow(\''+key+'\')" style="cursor:pointer"><td style="color:var(--mu)">'+liveArrow_(key)+(i+1)+'</td><td>'+eh(r.tradingsymbol)+'</td><td style="color:var(--dm);font-size:10px">'+eh(r.product)+'</td>'+
      '<td>'+r.quantity+'</td><td>'+fNZ(r.average_price)+'</td><td>'+fNZ(r.last_price)+'</td>'+
      '<td class="'+pnlCls+'">'+(r.pnl>=0?"+":"")+"\u20B9"+Math.round(r.pnl).toLocaleString("en-IN")+'</td>'+liveActionCell_(r.tradingsymbol,intra)+'</tr>'+liveExpandRow_(key,r.tradingsymbol,intra,8);
  }).join("");
}
function renderLiveHoldings(rows){
  _liveHldAll=rows||[];_liveHldRows=_liveHldAll;updatePtFooter_(); // v2
  paintLiveHoldings_();
}
function paintLiveHoldings_(){
  var b=document.getElementById("liveHldTblBdy"),all=_liveHldAll;
  document.getElementById("liveHldCount").textContent=all.length;
  var q=liveQ_("liveHldSearch");
  var shown=all.filter(function(r){return !q||String(r.tradingsymbol).toUpperCase().indexOf(q)>=0;});
  var inv=0,cur=0,pnl=0,day=0,hasDay=false;
  all.forEach(function(r){var qn=Math.abs(Number(r.quantity)||0);inv+=qn*(Number(r.average_price)||0);cur+=qn*(Number(r.last_price)||0);pnl+=Number(r.pnl)||0;if(r.day_change!=null){day+=qn*(Number(r.day_change)||0);hasDay=true;}});
  document.getElementById("liveHldSummary").innerHTML=all.length?('Invested: <b>&#8377;'+fmtN_(inv)+'</b> &middot; Current: <b>&#8377;'+fmtN_(cur)+'</b> &middot; P&amp;L: <b class="'+(pnl>=0?"rp":"rn")+'">'+inr_(pnl)+'</b>'+(hasDay?' &middot; Today: <b class="'+(day>=0?"rp":"rn")+'">'+inr_(day)+'</b>':"")):"";
  if(!all.length){b.innerHTML='<tr class="emr"><td colspan="7">No holdings.</td></tr>';return;}
  if(!shown.length){b.innerHTML='<tr class="emr"><td colspan="7">No matching holdings.</td></tr>';return;}
  b.innerHTML=shown.map(function(r,i){
    var pnlCls=r.pnl>=0?"rp":"rn",key="h|"+String(r.tradingsymbol).toUpperCase();
    return '<tr onclick="toggleLiveRow(\''+key+'\')" style="cursor:pointer"><td style="color:var(--mu)">'+liveArrow_(key)+(i+1)+'</td><td>'+eh(r.tradingsymbol)+'</td>'+
      '<td>'+r.quantity+'</td><td>'+fNZ(r.average_price)+'</td><td>'+fNZ(r.last_price)+'</td>'+
      '<td class="'+pnlCls+'">'+(r.pnl>=0?"+":"")+"\u20B9"+Math.round(r.pnl).toLocaleString("en-IN")+'</td>'+liveActionCell_(r.tradingsymbol,false)+'</tr>'+liveExpandRow_(key,r.tradingsymbol,false,7);
  }).join("");
}
function loadLiveOrders(){ if(_obPanelOpen&&_obMode==="live")loadOrders(true); } // v4: live orders are shown in the Order Book
/* ── Live Trade Log (same renderers as Mock, fed with the broker's trades) ── */
function liveRowsForCtx_(){
  return _liveMirror.map(function(m){
    var intra=m.category==="intraday",src=intra?_livePosAll:_liveHldAll.concat(_livePosAll),ltp=null;
    for(var i=0;i<src.length;i++){if(String(src[i].tradingsymbol).toUpperCase()===m.symbol&&Number(src[i].last_price)>0){ltp=Number(src[i].last_price);break;}}
    return{symbol:m.symbol,category:m.category,qty:m.qty,avgBuyPrice:m.avgBuyPrice,tradeType:m.tradeType,ltp:ltp,action:m.action};
  });
}
function withLiveCtx_(fn){
  var sp=_ptRows,st=_allTradesBySymbol,sm=_netMemo,sx=_tlExpanded,sr=_tlReadOnly;
  _ptRows=liveRowsForCtx_();_allTradesBySymbol=_liveTradesBySymbol||{};_netMemo={};_tlExpanded=_liveTlExpanded;_tlReadOnly=true;
  try{return fn();}finally{_ptRows=sp;_allTradesBySymbol=st;_netMemo=sm;_tlExpanded=sx;_tlReadOnly=sr;}
}
function loadLiveTrades(){
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    _liveTradesBySymbol=(r&&r.success)?(r.bySymbol||{}):{};
    renderLiveTradeLog_();
  }).withFailureHandler(function(){}).getLiveTradesFromWeb(_authToken);
}
function renderLiveTradeLog_(){
  var b=document.getElementById("liveTlTblBdy");if(!b)return;
  if(_liveTradesBySymbol===null){b.innerHTML='<tr class="emr"><td colspan="10">Loading...</td></tr>';return;}
  withLiveCtx_(function(){
    var groups=tlGroupsAll_();
    document.getElementById("liveTlCount").textContent=groups.length;
    var sum=document.getElementById("liveTlSummary");
    if(!groups.length){b.innerHTML='<tr class="emr"><td colspan="10">No live trades recorded yet. Trades are captured from Zerodha as they happen, from today onward.</td></tr>';sum.innerHTML="";return;}
    var ts=tlSummarize_(groups),xv=computeGroupsXIRR_(groups);
    var xs=xv!=null?(" &middot; XIRR: <b class=\""+(xv>=0?"rp":"rn")+"\">"+(xv>=0?"+":"")+(xv*100).toFixed(1)+"%</b>"):"";
    var cg=catGainLine_(groups.map(function(g){return{category:g.category,gain:g.pnl||0};}));
    sum.innerHTML='Invested: <b>&#8377;'+Math.round(ts.invested).toLocaleString("en-IN")+'</b> &middot; Current/Exit Val.: <b>&#8377;'+Math.round(ts.current).toLocaleString("en-IN")+'</b> &middot; Gain/Loss: <b class="'+(ts.gain>=0?"rp":"rn")+'">'+(ts.gain>=0?"+":"")+'&#8377;'+Math.round(ts.gain).toLocaleString("en-IN")+'</b>'+chargesHtml_(ts)+xs+(cg?' &middot; '+cg:'');
    var q=liveQ_("liveTlSearch"),c=(document.getElementById("liveTlCatFilter")||{}).value||"";
    var shown=groups.filter(function(g){return rowMatches_(g,q,c);});
    if(!shown.length){b.innerHTML='<tr class="emr"><td colspan="10">No matching stocks.</td></tr>';return;}
    patchTbody(b,shown.map(function(g,i){
      var it=tlRowHtml_(g,i);
      it.rows=it.rows.map(function(h){return h.replace(/tlExp_/g,"ltlExp_").replace(/toggleTlRow\(/g,"toggleLiveTlRow(");});
      return it;
    }));
  });
}
function toggleLiveTlRow(symbol,category){var k=symbol+"__"+category;_liveTlExpanded[k]=!_liveTlExpanded[k];renderLiveTradeLog_();}

/* ── Live orders live in the ORDER BOOK (Mock | Live tabs) ── */
var _obMode="mock";
function obAvail_(){return _tradingAccountPref==="both"?"both":(_tradingAccountPref==="live"?"live":"mock");}
function applyObMode_(){
  var av=obAvail_();
  if(av==="live")_obMode="live";else if(av==="mock")_obMode="mock";
  var live=_obMode==="live";
  document.getElementById("obModeTabs").style.display=av==="both"?"inline-flex":"none";
  document.getElementById("obModeMock").style.cssText="border-color:"+(live?"var(--bd)":"var(--ac)")+";color:"+(live?"var(--dm)":"var(--ac)");
  document.getElementById("obModeLive").style.cssText="border-color:"+(live?"var(--rd)":"var(--bd)")+";color:"+(live?"var(--rd)":"var(--dm)");
  document.getElementById("obHeadMock").style.display=live?"none":"";document.getElementById("obHeadLive").style.display=live?"":"none";
  document.getElementById("obFilterStatus").style.display=live?"none":"";document.getElementById("obFilterCat").style.display=live?"none":"";
  document.getElementById("obFilterStatusLive").style.display=live?"":"none";document.getElementById("obFilterProdLive").style.display=live?"":"none";
  document.getElementById("obTtl").innerHTML="&#9654; Order Book"+(live?' <span style="color:var(--rd);font-size:10px;letter-spacing:1px">LIVE</span>':"");
}
function setObMode(m){_obMode=m;applyObMode_();loadOrders();}
function renderLiveOrders(rows){_liveOrderRows=rows||[];renderLiveOrderBook_();}
function renderLiveOrderBook_(){
  var b=document.getElementById("obTblBdy");if(!b)return;
  var all=_liveOrderRows;
  document.getElementById("obCount").textContent=all.length;
  var q=((document.getElementById("obSearch")||{}).value||"").trim().toUpperCase(),sf=(document.getElementById("obFilterStatusLive")||{}).value||"",pf=(document.getElementById("obFilterProdLive")||{}).value||"";
  var shown=all.filter(function(o){return(!q||String(o.symbol).toUpperCase().indexOf(q)>=0)&&(!sf||liveOrdClass_(o.status)===sf)&&(!pf||String(o.product).toUpperCase()===pf);});
  var items=[];
  if(!_liveOrderMeta.ok)items.push({key:"__banner__",rows:['<tr><td colspan="10" style="background:rgba(255,180,60,.12);color:var(--am);font-size:10px;padding:8px">'+(_liveOrderMeta.reconnect?"Zerodha is not connected today \u2014 showing only this app's own records (e.g. refusals). Connect in Portfolio \u2192 Live.":"Could not read the broker's order list ("+eh(_liveOrderMeta.err||"unknown")+") \u2014 showing this app's own records.")+'</td></tr>']});
  if(!shown.length){
    items.push({key:"__empty__",rows:['<tr class="emr"><td colspan="10">No live orders'+(all.length?' matching filter':' today')+'.</td></tr>']});
    patchTbody(b,items);return;
  }
  shown.forEach(function(o,i){
    var buy=String(o.side).toUpperCase()!=="SELL";
    var typeBdg=buy?'<span class="cbdg" style="background:rgba(70,200,140,.15);color:var(--gr)">BUY</span>':'<span class="cbdg" style="background:rgba(255,90,90,.15);color:var(--rd)">SELL</span>';
    var cls=liveOrdClass_(o.status),clr=cls==="COMPLETE"?"var(--gr)":(cls==="REJECTED"||cls==="CANCELLED")?"var(--rd)":"var(--am)";
    var dt=o.ts?new Date(o.ts.replace(" ","T")).toLocaleTimeString("en-IN",{hour12:false,hour:"2-digit",minute:"2-digit"}):"-";
    var prod=String(o.product||"").toUpperCase(),prodLbl=prod==="MIS"?"Intraday (MIS)":(prod==="CNC"?"Delivery (CNC)":eh(prod||"-"));
    var qtyTxt=(o.filled&&o.filled!==o.qty)?(o.filled+"/"+o.qty):String(o.qty);
    // v5: the reason is a note icon right next to the status (no separate column); open orders can be changed or cancelled
    var acts=(cls==="OPEN"&&o.orderId&&o.source==="broker")?('<button class="rmBtn" title="Change quantity / price" onclick="event.stopPropagation();modifyLiveOrder(\''+eh(o.orderId)+'\')">&#9998;</button><button class="rmBtn" title="Cancel this order" onclick="event.stopPropagation();cancelLiveOrder(\''+eh(o.orderId)+'\')">&#10005;</button>'):"";
    var html='<tr><td style="color:var(--mu)">'+(i+1)+'</td><td style="font-weight:700">'+eh(o.symbol)+(o.auto?' <span class="cbdg" style="background:rgba(77,166,255,.12);color:var(--bl)">AI</span>':'')+'</td>'+
      '<td style="color:var(--dm);font-size:10px">'+prodLbl+'</td><td>'+typeBdg+'</td><td>'+qtyTxt+'</td><td>'+(o.price?fNZ(o.price):"MKT")+'</td>'+
      '<td>\u20B9'+Math.round((o.qty||0)*(o.price||0)).toLocaleString("en-IN")+'</td>'+
      '<td style="color:'+clr+';font-weight:600">'+eh(o.status)+(o.reason?' '+noteIco(o.reason):'')+'</td><td style="color:var(--dm);font-size:10px">'+dt+'</td>'+
      '<td style="white-space:nowrap">'+acts+'</td></tr>';
    items.push({key:(o.orderId||"a")+"_"+i,rows:[html]});
  });
  patchTbody(b,items);
}
function modifyLiveOrder(id){
  var o=_liveOrderRows.filter(function(x){return x.orderId===id;})[0];if(!o)return;
  var q=prompt("New quantity for "+o.symbol+":",o.qty);if(q===null)return;
  var p=o.price;
  if(String(o.orderType).toUpperCase()==="LIMIT"||o.price){p=prompt("New limit price (\u20B9):",o.price);if(p===null)return;}
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){loadOrders(true);if(_tradingMode==="live")loadLivePortfolio();}
    else alert("Could not change the order: "+(r&&r.error?r.error:"unknown error")+(r&&r.needsReconnect?"\nOpen Portfolio \u2192 Live and press Connect Zerodha first.":""));
  }).withFailureHandler(function(e){alert("Could not change the order: "+(e&&e.message?e.message:"unknown error"));}).kiteModifyOrderFromWeb(_authToken,id,q,p);
}
function cancelLiveOrder(id){
  var o=_liveOrderRows.filter(function(x){return x.orderId===id;})[0];if(!o)return;
  if(!confirm("Cancel the pending "+o.side+" order for "+o.qty+" \u00D7 "+o.symbol+"?"))return;
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){loadOrders(true);}
    else alert("Could not cancel the order: "+(r&&r.error?r.error:"unknown error"));
  }).withFailureHandler(function(e){alert("Could not cancel the order: "+(e&&e.message?e.message:"unknown error"));}).kiteCancelOrderFromWeb(_authToken,id);
}
function liveOrdClass_(st){
  st=String(st||"").toUpperCase();
  if(st==="COMPLETE")return"COMPLETE";
  if(st==="REJECTED")return"REJECTED";
  if(st==="CANCELLED")return"CANCELLED";
  return"OPEN";
}
function paintLiveOrders_(){renderLiveOrderBook_();}
function paintLiveSub_(){
  if(_liveSubtab==="positions")paintLivePositions_();
  else if(_liveSubtab==="tradelog")renderLiveTradeLog_();
  else paintLiveHoldings_();
}
function exportLiveExcel_(){
  if(typeof XLSX==="undefined"){alert("The download library didn't load — check your connection and try again.");return;}
  var wb=XLSX.utils.book_new();
  var act=function(sym,intra){var m=liveMirrorMatch_(sym,intra);return m&&m.action?m.action:"";};
  var pos=[["Symbol","Product","Qty","Avg Price","LTP","P&L","Action"]].concat(_livePosAll.map(function(r){return[r.tradingsymbol,r.product,r.quantity,r.average_price,r.last_price,Math.round(r.pnl),act(r.tradingsymbol,String(r.product).toUpperCase()!=="CNC")];}));
  var hld=[["Symbol","Qty","Avg Price","LTP","P&L","Action"]].concat(_liveHldAll.map(function(r){return[r.tradingsymbol,r.quantity,r.average_price,r.last_price,Math.round(r.pnl),act(r.tradingsymbol,false)];}));
  var tl=[["Symbol","Category","Status","Net Qty","Invested","Current/Exit Val.","P&L","Return %"]];
  withLiveCtx_(function(){tlGroupsAll_().forEach(function(g){tl.push([g.symbol,g.category,g.isOpen?"Running":"Closed",g.netQty,Math.round(g.invested||0),Math.round(g.currentVal||0),g.pnl!=null?Math.round(g.pnl):"",g.retPct!=null?Math.round(g.retPct*10)/10:""]);});});
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(hld),"Holdings");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(pos),"Positions");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(tl),"Trade Log");
  XLSX.writeFile(wb,"live-report-"+new Date().toISOString().slice(0,10)+".xlsx");
}
function switchLiveSubtab(sub){
  _liveSubtab=sub;
  ["holdings","positions","tradelog"].forEach(function(s){
    document.getElementById("liveSub-"+s).style.display=(s===sub)?"":"none";
    document.getElementById("liveSubBtn-"+s).className="lptab"+(s===sub?" ptsub-on":"");
  });
  document.querySelectorAll("#liveSubTabsBar .ptCtl").forEach(function(el){el.style.display=(el.getAttribute("data-for")===sub)?"":"none";});
  updatePtFooter_(); // v2
  paintLiveSub_();
}
function placeLiveOrder(){
  var symbol=(document.getElementById("liveOrdSymbol").value||"").trim().toUpperCase();
  var category=document.getElementById("liveOrdCategory").value;
  var side=document.getElementById("liveOrdSide").value;
  var qty=parseInt(document.getElementById("liveOrdQty").value,10);
  var orderKind=document.getElementById("liveOrdType").value;
  var price=parseFloat(document.getElementById("liveOrdPrice").value)||null;
  if(!symbol||!qty||qty<1){alert("Enter a symbol and a quantity of at least 1.");return;}
  if(orderKind==="LIMIT"&&!price){alert("Enter a limit price, or switch order type to Market.");return;}
  // Catches a fat-fingered price (typing 2500 instead of 250, an extra zero, etc.) before it
  // becomes a real order — only checks when the picked symbol's LTP is actually known.
  if(orderKind==="LIMIT"&&price&&_liveOrdLtp&&Math.abs(price-_liveOrdLtp)/_liveOrdLtp>0.1){
    if(!confirm("Your limit price \u20B9"+price+" is more than 10% away from the last traded price \u20B9"+_liveOrdLtp+" for "+symbol+". Continue anyway?"))return;
  }
  var estValue=(price||_liveOrdLtp||0)*qty;
  if(!confirm("Place a REAL "+side+" order for "+qty+" "+symbol+" ("+orderKind+(price?" @ \u20B9"+price:"")+")"+(estValue?", approx \u20B9"+Math.round(estValue).toLocaleString("en-IN"):"")+"? This sends a live order to your broker."))return;
  var btn=document.getElementById("liveOrdSubmitBtn");
  if(btn){btn.disabled=true;btn.textContent="Sending\u2026";}
  google.script.run.withSuccessHandler(function(r){
    if(btn){btn.disabled=false;btn.textContent="Place Order";}
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){alert("Order sent to Zerodha (ID: "+r.orderId+"). If the broker or exchange rejects it a moment later, the reason will show in the Orders tab.");loadLiveOrders();loadLivePortfolio();setTimeout(loadLiveOrders,3000);setTimeout(loadLiveOrders,10000);}
    else alert(r&&r.error?r.error:"Order failed.");
  }).withFailureHandler(function(e){
    if(btn){btn.disabled=false;btn.textContent="Place Order";}
    alert("Error: "+(e&&e.message?e.message:"unknown"));
  }).kitePlaceOrderFromWeb(_authToken,symbol,category,side,qty,price,orderKind);
}
var _tlExpanded={};
function tlAllRows_(){
  var rows=[];
  Object.keys(_allTradesBySymbol||{}).forEach(function(sym){
    (_allTradesBySymbol[sym]||[]).forEach(function(t){rows.push(Object.assign({symbol:sym},t));});
  });
  rows.sort(function(a,b){return new Date(b.tradeDate||0)-new Date(a.tradeDate||0);});
  return rows;
}
function tlGroupStats_(sym,trades,category){
  var totalBuyCost=0,totalSellProceeds=0,netQty=0;
  trades.forEach(function(t){
    if(t.direction==="SELL"){totalSellProceeds+=t.qty*t.price;netQty-=t.qty;}
    else{totalBuyCost+=t.qty*t.price;netQty+=t.qty;}
  });
  var isOpen=netQty>0;
  var posRow=isOpen?_ptRows.find(function(p){return p.symbol===sym&&p.category===category;}):null;
  // v2: P&L is realized + unrealized from the same FIFO pass that produces the per-trade lines, net of brokerage &
  // charges when enabled - so this header always equals the sum of its own lines. (A running group used to show only
  // the unrealized part and silently drop profit already banked by partial sells.)
  var fa=fifoAnalyze_(trades,posRow?posRow.ltp:null,category||"longterm"),T=fa.totals;
  var matchedBase=0;fa.lines.forEach(function(L){matchedBase+=L.matchedBase;});
  var invested,currentVal,pnl,retPct;
  if(isOpen&&posRow){
    invested=posRow.avgBuyPrice*posRow.qty;
    currentVal=posRow.ltp!=null?posRow.ltp*posRow.qty:null;
    pnl=T.unrealized!=null?(T.realized+T.unrealized):null;
    var base=T.openCost+matchedBase;
    retPct=(pnl!=null&&base)?pnl/base*100:null;
  } else {
    invested=totalBuyCost;
    currentVal=totalSellProceeds;
    pnl=T.realized+(T.unrealized||0);
    retPct=totalBuyCost?pnl/totalBuyCost*100:null;
  }
  var xirrVal=computeStockXIRR(trades,posRow?posRow.ltp:null);
  var lastDate=trades.reduce(function(m,t){var d=new Date(t.tradeDate||0);return d>m?d:m;},new Date(0));
  return {symbol:sym,category:category||"longterm",netQty:netQty,isOpen:isOpen,invested:invested,currentVal:currentVal,pnl:pnl,retPct:retPct,xirr:xirrVal,trades:trades,lastDate:lastDate,
          realized:T.realized,unrealized:T.unrealized,charges:T.charges};
}
function tlGroupsAll_(){
  var groups=[];
  Object.keys(_allTradesBySymbol||{}).forEach(function(sym){
    var bySymCategory={};
    (_allTradesBySymbol[sym]||[]).forEach(function(t){
      var cat=t.category||"longterm";
      (bySymCategory[cat]=bySymCategory[cat]||[]).push(t);
    });
    // One group per symbol+category — a stock traded in more than one category (e.g. both intraday
    // and long-term) is two separate rows here, not one row blending both together. This was
    // previously grouping by symbol alone, which is what showed as merged categories.
    Object.keys(bySymCategory).forEach(function(cat){
      groups.push(tlGroupStats_(sym,bySymCategory[cat],cat));
    });
  });
  groups.sort(function(a,b){return b.lastDate-a.lastDate;});
  return groups;
}
function computeGroupsXIRR_(groups){
  if(!groups||!groups.length)return null;
  var cfs=[],dates=[];
  groups.forEach(function(g){
    if(!g.trades||!g.trades.length||!g.invested)return;
    var firstDate=g.trades.reduce(function(m,t){var d=new Date(t.tradeDate||0);return(!isNaN(d)&&d<m)?d:m;},new Date(8640000000000000));
    cfs.push(-g.invested);dates.push(firstDate);
    if(g.currentVal!=null)cfs.push(g.currentVal),dates.push(g.isOpen?new Date():g.lastDate);
  });
  if(cfs.length<2)return null;
  return xirr(cfs,dates);
}
function tlSummarize_(groups){
  var invested=0,current=0,gain=0,charges=0;
  groups.forEach(function(g){invested+=g.invested||0;current+=(g.currentVal!=null?g.currentVal:g.invested)||0;gain+=(g.pnl!=null?g.pnl:0);charges+=g.charges||0;});
  return {invested:invested,current:current,gain:gain,charges:charges};
}
function toggleTlRow(symbol,category){
  var pk=symbol+"__"+category;
  _tlExpanded[pk]=!_tlExpanded[pk];
  var row=document.getElementById("tlExp_"+pk);
  if(row)row.style.display=_tlExpanded[pk]?"table-row":"none";
}
function renderTradeLog(){
  var b=document.getElementById("tlTblBdy");if(!b)return;
  if(_allTradesBySymbol===null){b.innerHTML='<tr class="emr"><td colspan="10">Loading...</td></tr>';loadPortfolioXirr_();return;}
  var groups=tlGroupsAll_();
  document.getElementById("tlCount").textContent=groups.length;
  var sumEl=document.getElementById("tlSummary");
  if(!groups.length){b.innerHTML='<tr class="emr"><td colspan="10">No trades yet.</td></tr>';if(sumEl)sumEl.innerHTML="";return;}
  var ts=tlSummarize_(groups);
  var xirrVal=computeGroupsXIRR_(groups);
  var xirrStr=xirrVal!=null?(" &middot; XIRR: <b class=\""+(xirrVal>=0?"rp":"rn")+"\">"+(xirrVal>=0?"+":"")+(xirrVal*100).toFixed(1)+"%</b>"):"";
  // v2: Gain/Loss by timeframe in the header (as Positions already did), net of charges when enabled
  var tCat=catGainLine_(groups.map(function(g){return{category:g.category,gain:g.pnl||0};}));
  if(sumEl)sumEl.innerHTML='Invested: <b>&#8377;'+Math.round(ts.invested).toLocaleString("en-IN")+'</b> &middot; '+
    'Current/Exit Val.: <b>&#8377;'+Math.round(ts.current).toLocaleString("en-IN")+'</b> &middot; '+
    'Gain/Loss: <b class="'+(ts.gain>=0?"rp":"rn")+'">'+(ts.gain>=0?"+":"")+'&#8377;'+Math.round(ts.gain).toLocaleString("en-IN")+'</b>'+chargesHtml_(ts)+xirrStr+
    (tCat?' &middot; '+tCat:'');
  var q=((document.getElementById("tlSearch")||{}).value||"").trim().toUpperCase(),cat=catFilterVal_("tlCatFilter");
  var shown=groups.filter(function(g){return rowMatches_(g,q,cat);});
  if(!shown.length){b.innerHTML='<tr class="emr"><td colspan="10">No matching stocks.</td></tr>';return;}
  patchTbody(b,shown.map(function(g,i){return tlRowHtml_(g,i);}));
}
function filterTradeLog(){renderTradeLog();} // v2: renderTradeLog now applies the search + category filters itself (so a background refresh no longer clears them)
function tlRowHtml_(g,i){
  var pk=g.symbol+"__"+g.category;
  var catLabel={intraday:"Intraday",shortterm:"Short-term",longterm:"Long-term"}[g.category]||g.category||"-";
  var statusBdg=g.isOpen?'<span class="cbdg" style="background:rgba(255,180,60,.15);color:var(--am)">Running</span>':(g.pnl!=null&&g.pnl>=0?'<span class="cbdg" style="background:rgba(70,200,140,.15);color:var(--gr)">Success</span>':'<span class="cbdg" style="background:rgba(255,90,90,.15);color:var(--rd)">Failed</span>');
  var pnlCls=g.pnl!=null?(g.pnl>=0?"rp":"rn"):"";
  var pnlStr=g.pnl!=null?((g.pnl>=0?"+":"")+"\u20B9"+Math.round(g.pnl).toLocaleString("en-IN")):"-";
  var retStr=g.retPct!=null?((g.retPct>=0?"+":"")+g.retPct.toFixed(1)+"%"):"-";
  var xirrStr=g.xirr!=null?((g.xirr>=0?"+":"")+(g.xirr*100).toFixed(1)+"%"):"-";
  var xirrCls=g.xirr!=null?(g.xirr>=0?"rp":"rn"):"";
  var expIco='<span class="ptExp" onclick="event.stopPropagation();toggleTlRow(\''+eh(g.symbol)+'\',\''+eh(g.category)+'\')" style="cursor:pointer;display:inline-block;width:14px;transform:'+(_tlExpanded[pk]?"rotate(90deg)":"rotate(0deg)")+';transition:transform .15s">&#9656;</span>';
  var mainRow='<tr onclick="toggleTlRow(\''+eh(g.symbol)+'\',\''+eh(g.category)+'\')" style="cursor:pointer"><td style="color:var(--mu)">'+expIco+(i+1)+'</td>'+symCell(g.symbol,"",null,g.category)+'<td>'+statusBdg+'</td>'+
    '<td style="color:var(--dm)">'+eh(catLabel)+'</td><td>'+g.netQty+'</td>'+
    '<td>'+(g.invested!=null?Math.round(g.invested).toLocaleString("en-IN"):"-")+'</td>'+
    '<td>'+(g.currentVal!=null?Math.round(g.currentVal).toLocaleString("en-IN"):"-")+'</td>'+
    '<td class="'+pnlCls+'"'+grpTitle_(g)+'>'+pnlStr+'</td><td class="'+pnlCls+'">'+retStr+'</td>'+
    '<td class="'+xirrCls+'">'+xirrStr+'</td></tr>';
  var expRow='<tr id="tlExp_'+eh(pk)+'" style="display:'+(_tlExpanded[pk]?"table-row":"none")+'"><td colspan="10" style="padding:0;background:var(--sf2)"><div style="padding:10px 16px;font-size:11px;color:var(--dm)">'+buildTradesHtml(g.symbol,g.trades,g.category)+'</div></td></tr>';
  return {key:pk,rows:[mainRow,expRow]};
}
function toggleTradeRow(symbol,category,ns){
  var pre=(ns==="pos")?"pos":"pt";
  var pk=symbol+"__"+category,ek=(ns==="pos"?"pos|":"")+pk;
  _ptExpanded[ek]=!_ptExpanded[ek];
  var row=document.getElementById(pre+"Exp_"+pk);
  if(!row)return;
  row.style.display=_ptExpanded[ek]?"table-row":"none";
  document.querySelectorAll(".ptExp").forEach(function(ic){if(ic.getAttribute("data-ek")===ek)ic.style.transform=_ptExpanded[ek]?"rotate(90deg)":"rotate(0deg)";});
  if(_ptExpanded[ek])loadTradesFor(symbol,category);
}
function loadTradesFor(symbol,category,silent){
  var pk=symbol+"__"+category;
  var el=document.getElementById("ptTrades_"+pk),el2=document.getElementById("posTrades_"+pk);
  if(!el&&!el2)return;
  if(!silent&&!_ptTradesCache[pk]){if(el)el.innerHTML="Loading trades...";if(el2)el2.innerHTML="Loading trades...";}
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(!r||!r.success){if(!silent){var m="Could not load trades: "+eh(r&&r.error?r.error:"unknown");if(el)el.innerHTML=m;if(el2)el2.innerHTML=m;}return;}
    _ptTradesCache[pk]=r.trades||[];
    renderTradesFor(symbol,category,r.trades||[]);
  }).withFailureHandler(function(e){if(!silent){var m="Server error: "+eh(e&&e.message?e.message:"unknown");if(el)el.innerHTML=m;if(el2)el2.innerHTML=m;}}).getPortfolioTradesFromWeb(_authToken,symbol,category);
}
function buildTradesHtml(symbol,trades,category,opts){
  if(!trades.length)return"No trades found.";
  opts=opts||{};
  var posRow=(_ptRows||[]).find(function(x){return x.symbol===symbol&&x.category===category;});
  var ltp=posRow?posRow.ltp:null;
  // v2: P&L per line now comes from a first-in-first-out pass over the trades (see fifoAnalyze_). Realized profit is
  // booked ONCE, on the trade that closed the lots (a SELL that closes a long, a BUY that covers a short), measured
  // against the actual lots it closed; an opening trade only shows P&L for whatever quantity is still open (marked at
  // LTP). The old logic measured every SELL against one stale average and every BUY against the average exit, so a
  // single round trip was counted on both lines and the lines never added up to the position's real P&L.
  var fa=fifoAnalyze_(trades,ltp,category);
  var shown=opts.todayOnly?trades.filter(function(t){return isTodayLocal_(t.tradeDate);}):trades;
  var rows=shown.map(function(t){
    var L=fa.lines.get(t)||{};
    var typeBdg=t.direction==="SELL"?'<span class="cbdg" style="background:rgba(255,90,90,.15);color:var(--rd)">SELL</span>':'<span class="cbdg" style="background:rgba(70,200,140,.15);color:var(--gr)">BUY</span>';
    var dt=t.tradeDate?new Date(t.tradeDate).toLocaleDateString("en-IN")+' '+new Date(t.tradeDate).toLocaleTimeString("en-IN",{hour12:false,hour:"2-digit",minute:"2-digit"}):"-";
    var amount=t.qty*t.price;
    var pnlCell='<span style="color:var(--dm)" title="Fully closed - its profit or loss is booked on the sell/cover line(s) that closed it">-</span>',retCell="-",xirrCell="-";
    if(L.pnl!=null){
      var rc=L.pnl>=0?"rp":"rn";
      pnlCell='<span class="'+rc+'" title="'+eh(lineTip_(L))+'">'+(L.pnl>=0?"+":"-")+"\u20B9"+Math.round(Math.abs(L.pnl)).toLocaleString("en-IN")+'</span>';
      var ret=L.base>0?L.pnl/L.base:null;
      retCell=ret!=null?'<span class="'+rc+'">'+(ret>=0?"+":"")+(ret*100).toFixed(1)+'%</span>':"-";
    } else if(L.openQty>0){
      pnlCell="-"; // still open but no live price yet
    }
    var lx=lineXirr_(L,t,ltp);
    if(lx!=null)xirrCell=(lx>=0?"+":"")+(lx*100).toFixed(1)+"%";
    var noteCell=noteIco(t.notes)||'<span style="color:var(--dm)">-</span>';
    return '<tr><td>'+typeBdg+'</td><td>'+t.qty+'</td><td>'+fNZ(t.price)+'</td>'+
      '<td>\u20B9'+Math.round(amount).toLocaleString("en-IN")+'</td>'+
      '<td>'+pnlCell+'</td><td>'+retCell+'</td><td>'+xirrCell+'</td>'+
      '<td style="color:var(--dm)">'+dt+'</td>'+
      '<td style="color:var(--dm);font-size:10px">'+eh(t.source||"Manual")+'</td>'+
      '<td>'+noteCell+'</td>'+
      '<td style="white-space:nowrap">'+(_tlReadOnly?'':'<button class="rmBtn" title="Edit" onclick="editTrade(\''+t.tradeId+'\',\''+symbol+'\','+t.qty+','+t.price+',\''+t.tradeDate+'\',\''+t.direction+'\',\''+t.category+'\')">&#9998;</button>'+
      '<button class="rmBtn" title="Delete" onclick="deleteTrade(\''+t.tradeId+'\',\''+symbol+'\',\''+category+'\')">&#10005;</button>')+'</td></tr>';
  }).join("");
  return '<table style="width:100%"><thead><tr><th>Type</th><th>Qty</th><th>Price</th><th>Amount</th><th>P&amp;L</th><th>Return</th><th>XIRR</th><th>Date</th><th>Source</th><th>Notes</th><th></th></tr></thead><tbody>'+rows+'</tbody></table>';
}
function renderTradesFor(symbol,category,trades){
  var pk=symbol+"__"+category;
  var el=document.getElementById("ptTrades_"+pk),el2=document.getElementById("posTrades_"+pk);
  if(!el&&!el2)return;
  if(el)el.innerHTML=buildTradesHtml(symbol,trades,category);
  if(el2){
    var pr=splitPtRows_().positions.find(function(x){return x.symbol===symbol&&x.category===category;})||{};
    el2.innerHTML=buildTradesHtml(symbol,trades,category,{todayOnly:!!pr._slice}); // a "traded today" row lists only today's trades
  }
  syncStickyOffsets_();
}
function computeStockXIRR(trades,ltp){
  if(!trades||!trades.length)return null;
  var cfs=[],dates=[],netQty=0;
  trades.slice().sort(function(a,b){return new Date(a.tradeDate)-new Date(b.tradeDate);}).forEach(function(t){
    var d=new Date(t.tradeDate);if(isNaN(d))d=new Date();
    var dir=t.direction==="SELL"?-1:1;
    cfs.push(-t.qty*t.price*dir);dates.push(d);
    netQty+=t.qty*dir;
  });
  if(netQty!==0){
    // Still open: value the remaining position with a hypothetical "sell today at LTP" cashflow.
    if(!ltp)return null;
    var currentVal=Math.abs(netQty)*ltp*(netQty>0?1:-1);
    cfs.push(currentVal);dates.push(new Date());
  }
  // netQty===0: fully closed, the cashflows above are already the complete realized
  // sequence — no synthetic exit needed, XIRR can be computed directly from them.
  if(cfs.length<2)return null;
  return xirr(cfs,dates);
}
function editTrade(tradeId,symbol,qty,price,date,direction,category){
  var newQty=prompt("Quantity:",qty);if(newQty===null)return;
  var newPrice=prompt("Price:",price);if(newPrice===null)return;
  var newDate=prompt("Date (YYYY-MM-DD):",date?date.slice(0,10):"");if(newDate===null)return;
  var newDir=prompt("Direction (BUY or SELL):",direction||"BUY");if(newDir===null)return;
  newDir=String(newDir).trim().toUpperCase()==="SELL"?"SELL":"BUY";
  var newCat=prompt("Category (intraday / shortterm / longterm):",category||"longterm");if(newCat===null)return;
  newCat=["intraday","shortterm","longterm"].indexOf(newCat.trim().toLowerCase())>=0?newCat.trim().toLowerCase():"longterm";
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){loadTradesFor(symbol,category);loadPT(true);}
    else alert(r&&r.error?r.error:"Could not update trade.");
  }).withFailureHandler(function(e){alert("Error: "+(e&&e.message?e.message:"unknown"));}).updatePortfolioTradeFromWeb(_authToken,tradeId,newQty,newPrice,newDate,"",newDir,newCat);
}
function deleteTrade(tradeId,symbol,category){
  if(!confirm("Delete this trade? The position will be recalculated."))return;
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){loadTradesFor(symbol,category);loadPT(true);}
    else alert(r&&r.error?r.error:"Could not delete trade.");
  }).withFailureHandler(function(e){alert("Error: "+(e&&e.message?e.message:"unknown"));}).removePortfolioTradeFromWeb(_authToken,tradeId);
}
function markExecuted(orderId){
  if(!confirm("Mark this order as Executed? It will be added to your Portfolio automatically."))return;
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success)loadOrders(true);
    else alert(r&&r.error?r.error:"Failed to mark as executed.");
  }).withFailureHandler(function(e){alert("Error: "+(e&&e.message?e.message:"unknown"));}).markOrderExecutedFromWeb(_authToken,orderId);
}

/* ══ PROFILE ══ */
function openProfile(){
  document.getElementById("profM").classList.add("open");
  document.getElementById("profErr").classList.remove("show");document.getElementById("profMsg").classList.remove("show");
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){
      document.getElementById("profName").value=r.displayName;document.getElementById("profEmail").value=r.email;
      document.getElementById("profMargin").value=r.margin||"";
      document.getElementById("profInvested").textContent="\u20B9"+Math.round(r.invested||0).toLocaleString("en-IN");
      document.getElementById("profCap").value=r.perTradeCap||"";
      document.getElementById("autoIntraday").checked=!!r.autoIntraday;
      document.getElementById("autoShortterm").checked=!!r.autoShortterm;
      document.getElementById("autoLongterm").checked=!!r.autoLongterm;
      var ci=Number(r.contribPctIntraday)||0,cs=Number(r.contribPctShortterm)||0,cl=Number(r.contribPctLongterm)||0;
      document.getElementById("profContribIntraday").value=ci;document.getElementById("contribIntradayLbl").textContent=fmtContribPct_(ci);
      document.getElementById("profContribShortterm").value=cs;document.getElementById("contribShorttermLbl").textContent=fmtContribPct_(cs);
      document.getElementById("profContribLongterm").value=cl;document.getElementById("contribLongtermLbl").textContent=fmtContribPct_(cl);
      document.getElementById("holdShortOn").checked=!!r.holdPeriodShortOn;
      var hsw=r.holdPeriodShortWeeks||2;document.getElementById("holdShortWeeks").value=hsw;document.getElementById("holdShortLbl").textContent=hsw+"wk";
      document.getElementById("holdLongOn").checked=!!r.holdPeriodLongOn;
      var hlw=r.holdPeriodLongWeeks||16;document.getElementById("holdLongWeeks").value=hlw;document.getElementById("holdLongLbl").textContent=fmtHoldWeeks_(hlw);
      setExecModeUI_(r.executionMode||"manual");
      setTradingAccountPrefUI_(r.tradingAccountPref||_tradingAccountPref||"mock");
      fillProfileV2_(r); // v2: margin used/available, selection limits, brokerage, two-step verification, broker
    }
  }).withFailureHandler(function(){}).getProfileFromWeb(_authToken);
}
function fmtHoldWeeks_(w){w=Number(w)||1;var mo=Math.round(w/4.345);return w+"wk"+(mo>=1?" (~"+mo+"mo)":"");}
// Measures each frozen table's own header height into --outerThH on its scroll container, so the
// nested dropdown header can stick directly beneath it (see the .scrollTbl rules in CSS). Measured
// rather than hard-coded because the header's height shifts with font size, zoom and wrapping — a
// fixed pixel guess leaves either a gap or an overlap on anything but one exact setup.
function syncStickyOffsets_(){
  document.querySelectorAll(".scrollTbl").forEach(function(c){
    var th=c.querySelector(":scope>table>thead");
    if(th&&th.offsetHeight)c.style.setProperty("--outerThH",th.offsetHeight+"px");
  });
}
window.addEventListener("resize",syncStickyOffsets_);
// 0 isn't "zero percent of margin" — it's the documented opt-out that falls back to the flat
// per-trade cap, so the slider says "Cap" there rather than a misleading "0%".
function fmtContribPct_(v){v=Number(v)||0;return v===0?"Cap":v+"%";}
function setExecModeUI_(mode){
  document.querySelectorAll("#execModeBtns .tybtn").forEach(function(b){b.classList.toggle("sel",b.dataset.mode===mode);});
}
function setExecMode(mode){
  setExecModeUI_(mode);
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){showErr("profMsg","Execution mode set to "+r.executionMode+".");document.getElementById("profMsg").classList.add("show");document.getElementById("profErr").classList.remove("show");}
    else{showErr("profErr",(r&&r.error)||"Could not update execution mode.");document.getElementById("profErr").classList.add("show");}
  }).withFailureHandler(function(){}).setExecutionModeFromWeb(_authToken,mode);
}
function setTradingAccountPrefUI_(pref){
  document.querySelectorAll("#tradingAcctBtns .tybtn").forEach(function(b){b.classList.toggle("sel",b.dataset.acct===pref);});
}
function setTradingAccountPref(pref){
  setTradingAccountPrefUI_(pref);
  _tradingAccountPref=pref;
  applyTradingAccountPref_();
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){showErr("profMsg","Trading account set to "+r.tradingAccountPref+".");document.getElementById("profMsg").classList.add("show");document.getElementById("profErr").classList.remove("show");}
    else{showErr("profErr",(r&&r.error)||"Could not update trading account preference.");document.getElementById("profErr").classList.add("show");}
  }).withFailureHandler(function(){}).setTradingAccountPrefFromWeb(_authToken,pref);
}
function saveMargin(){
  var amt=document.getElementById("profMargin").value;
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){showErr("profMsg","Margin updated.");loadPT(true);document.getElementById("profMsg").classList.add("show");document.getElementById("profErr").classList.remove("show");}
    else{showErr("profErr",(r&&r.error)||"Could not update margin.");document.getElementById("profErr").classList.add("show");}
  }).withFailureHandler(function(){}).setMarginFromWeb(_authToken,amt);
}
function saveCap(){
  var amt=document.getElementById("profCap").value;
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){showErr("profMsg","Per-trade cap updated.");document.getElementById("profErr").classList.remove("show");}
    else{showErr("profErr",(r&&r.error)||"Could not update cap.");document.getElementById("profMsg").classList.remove("show");}
  }).withFailureHandler(function(){}).setPerTradeCapFromWeb(_authToken,amt);
}
function saveAutoToggles(){
  var i=document.getElementById("autoIntraday").checked,s=document.getElementById("autoShortterm").checked,l=document.getElementById("autoLongterm").checked;
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){showErr("profMsg","Auto-trade preferences saved.");document.getElementById("profMsg").classList.add("show");document.getElementById("profErr").classList.remove("show");}
    else{showErr("profErr",(r&&r.error)||"Could not save preferences.");document.getElementById("profErr").classList.add("show");}
  }).withFailureHandler(function(){}).setAutoTradeTogglesFromWeb(_authToken,i,s,l);
}
function saveContribution(){
  var i=document.getElementById("profContribIntraday").value||0,s=document.getElementById("profContribShortterm").value||0,l=document.getElementById("profContribLongterm").value||0;
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){showErr("profMsg","Contribution % saved.");document.getElementById("profMsg").classList.add("show");document.getElementById("profErr").classList.remove("show");}
    else{showErr("profErr",(r&&r.error)||"Could not save contribution %.");document.getElementById("profErr").classList.add("show");}
  }).withFailureHandler(function(){}).setContributionPctFromWeb(_authToken,i,s,l);
}
function saveHoldPeriod(){
  var so=document.getElementById("holdShortOn").checked,sw=document.getElementById("holdShortWeeks").value;
  var lo=document.getElementById("holdLongOn").checked,lw=document.getElementById("holdLongWeeks").value;
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){showErr("profMsg","Hold period saved.");document.getElementById("profMsg").classList.add("show");document.getElementById("profErr").classList.remove("show");}
    else{showErr("profErr",(r&&r.error)||"Could not save hold period.");document.getElementById("profErr").classList.add("show");}
  }).withFailureHandler(function(){}).setHoldPeriodFromWeb(_authToken,so,sw,lo,lw);
}
function closeProfile(){document.getElementById("profM").classList.remove("open");}

/* ══ v2 SETTINGS: stock-selection limits · Brokerage & margin privilege · Broker (Live) · two-step verification ══ */
function setMsg_(msg){showErr("profMsg",msg);document.getElementById("profMsg").classList.add("show");document.getElementById("profErr").classList.remove("show");}
function setErr_(msg){showErr("profErr",msg);document.getElementById("profErr").classList.add("show");document.getElementById("profMsg").classList.remove("show");}
function fillProfileV2_(r){
  var mu=document.getElementById("profMarginUsed");
  if(mu){
    if(r.margin>0)mu.innerHTML='Margin used: <b style="color:var(--tx)">&#8377;'+fmtN_(r.usedMargin)+'</b> &middot; Available: <b style="color:var(--tx)">&#8377;'+fmtN_(r.availableMargin)+'</b>'+(r.leverage>1?' &middot; intraday leverage <b style="color:var(--tx)">'+r.leverage+'x</b>':'');
    else mu.textContent="";
  }
  var s=r.selection||{};
  document.getElementById("selPriceMin").value=s.priceMin||"";document.getElementById("selPriceMax").value=s.priceMax||"";
  document.getElementById("selUnitsMin").value=s.unitsMin||"";document.getElementById("selUnitsMax").value=s.unitsMax||"";
  document.getElementById("ntdOn").checked=!!s.tradeAnyDay;setNtdBanner_(!!s.tradeAnyDay);
  document.getElementById("liveAutoOn").checked=!!s.liveAuto;
  if(r.brokerage){_brk=r.brokerage;fillBrokerageForm_(r.brokerage);}
  mfaRender_(!!r.mfaEnabled);
  loadBrokerSettings_();
}
// v5: Live auto-trading is a per-user choice. OFF = the AI never touches the Live account (view / place live orders yourself).
function saveLiveAuto(){
  var cb=document.getElementById("liveAutoOn"),on=cb.checked,code="";
  if(on){
    if(!confirm("Let the AI place REAL orders (entries and exits) on your Zerodha account?\n\nIt only does so when Trading Account is Live or Both, execution mode is Auto, and your broker is connected today. You can switch this off at any time."))  {cb.checked=false;return;}
    if(_bkState&&_bkState.mfaEnabled){code=prompt("Two-step verification: enter the 6-digit code from your authenticator app.");if(code===null){cb.checked=false;return;}}
  }
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success)setMsg_(r.selection.liveAuto?"Live auto-trading is ON for your account.":"Live auto-trading is OFF. The AI will not place real orders.");
    else{cb.checked=!on;setErr_((r&&r.error)||"Could not save this setting.");}
  }).withFailureHandler(function(){cb.checked=!on;setErr_("Could not reach the server.");}).saveLiveAutoFromWeb(_authToken,on,code);
}
var _ntdOn=false;
function setNtdBanner_(on){_ntdOn=!!on;var b=document.getElementById("ntdBanner");if(b)b.style.display=on?"":"none";}
function saveNonTradingDay(){
  var cb=document.getElementById("ntdOn"),on=cb.checked;
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){setNtdBanner_(!!r.selection.tradeAnyDay);setMsg_(r.selection.tradeAnyDay?"Weekend & holiday trading is ON. Switch it off after the special session.":"Weekend & holiday trading is OFF.");}
    else{cb.checked=!on;setErr_((r&&r.error)||"Could not save this setting.");}
  }).withFailureHandler(function(){cb.checked=!on;setErr_("Could not reach the server.");}).saveNonTradingDayTradingFromWeb(_authToken,on);
}
// Loads the switch + charge settings once at start-up so the header banner (and net-of-charges P&L) are right on every tab.
function loadNtdFlag_(){
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.success){setNtdBanner_(!!(r.selection&&r.selection.tradeAnyDay));if(r.brokerage&&!_brk)_brk=r.brokerage;}
  }).withFailureHandler(function(){}).getBrokerageSettingsFromWeb(_authToken);
}

// ---- Broker (Live) setup guide ----
function openBrokerGuide(){document.getElementById("guideM").classList.add("open");loadGuideInfo_();}
function closeBrokerGuide(){document.getElementById("guideM").classList.remove("open");}
var _guideInfoLoaded=false;
function loadGuideInfo_(){
  if(_guideInfoLoaded)return;
  google.script.run.withSuccessHandler(function(r){
    var el=document.getElementById("gdRedirect");
    if(r&&r.success&&r.redirectUrl){el.textContent=r.redirectUrl;_guideInfoLoaded=true;}
    else el.textContent="(could not read it \u2014 use your dashboard's web-app address, the one ending in /exec)";
  }).withFailureHandler(function(){document.getElementById("gdRedirect").textContent="(could not read it \u2014 use your dashboard's web-app address, the one ending in /exec)";}).getBrokerGuideInfoFromWeb(_authToken);
}
function copyGuideBlock(btn,id){var t=document.getElementById(id).textContent;try{navigator.clipboard.writeText(t).then(function(){flashBtn_(btn,"Copied \u2713");});}catch(e){window.prompt("Copy:",t);}}
function downloadGuideBlock(btn,id,filename){
  var t=document.getElementById(id).textContent;
  try{var bl=new Blob([t],{type:"text/plain"}),u=URL.createObjectURL(bl),a=document.createElement("a");a.href=u;a.download=filename;document.body.appendChild(a);a.click();setTimeout(function(){document.body.removeChild(a);URL.revokeObjectURL(u);},600);flashBtn_(btn,"Saved \u2713");}
  catch(e){window.prompt("Copy the contents of "+filename+":",t);}
}
function copyGuideRedirect(btn){var t=document.getElementById("gdRedirect").textContent;try{navigator.clipboard.writeText(t).then(function(){flashBtn_(btn,"Copied \u2713");});}catch(e){window.prompt("Copy:",t);}}
function saveSelectionLimits(){
  var v=function(id){return document.getElementById(id).value;};
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success)setMsg_("Stock selection limits saved.");else setErr_((r&&r.error)||"Could not save selection limits.");
  }).withFailureHandler(function(){setErr_("Could not reach the server.");}).saveSelectionLimitsFromWeb(_authToken,v("selPriceMin"),v("selPriceMax"),v("selUnitsMin"),v("selUnitsMax"));
}

/* ── Brokerage ── */
var BRK_ROWS_=[["brokeragePct","Brokerage (% of order value)"],["brokerageCap","Brokerage cap (\u20B9 per executed order, 0 = none)"],["sttBuyPct","STT/CTT % \u2014 buy side"],["sttSellPct","STT/CTT % \u2014 sell side"],["txnNsePct","Transaction charges % \u2014 NSE"],["txnBsePct","Transaction charges % \u2014 BSE"],["gstPct","GST % (on brokerage + SEBI + transaction)"],["sebiPerCrore","SEBI charges (\u20B9 per crore)"],["ipftPerCrore","IPFT charges, NSE (\u20B9 per crore)"],["stampBuyPct","Stamp charges % \u2014 buy side"]];
var BRK_DEFAULTS_={broker:"Zerodha",applyToPnl:true,exchange:"NSE",
  delivery:{brokeragePct:0,brokerageCap:0,sttBuyPct:0.1,sttSellPct:0.1,txnNsePct:0.00307,txnBsePct:0.00375,gstPct:18,sebiPerCrore:10,ipftPerCrore:0.01,stampBuyPct:0.015},
  intraday:{brokeragePct:0.03,brokerageCap:20,sttBuyPct:0,sttSellPct:0.025,txnNsePct:0.00307,txnBsePct:0.00375,gstPct:18,sebiPerCrore:10,ipftPerCrore:0.01,stampBuyPct:0.003},
  marginPrivilege:{on:false,intradayLeverage:5}};
function buildBrokerageGrid_(){
  var el=document.getElementById("brkGrid");if(!el||el.firstChild)return;
  var h='<table class="chgTbl"><thead><tr><th></th><th>Equity delivery</th><th>Equity intraday</th></tr></thead><tbody>';
  BRK_ROWS_.forEach(function(r){h+='<tr><td>'+r[1]+'</td><td><input class="minp chgIn" type="number" step="any" min="0" id="brkD_'+r[0]+'"/></td><td><input class="minp chgIn" type="number" step="any" min="0" id="brkI_'+r[0]+'"/></td></tr>';});
  el.innerHTML=h+'</tbody></table>';
}
function fillBrokerageForm_(c){
  buildBrokerageGrid_();
  c=c||BRK_DEFAULTS_;
  document.getElementById("brkName").value=c.broker||"";
  document.getElementById("brkApply").checked=!!c.applyToPnl;
  document.getElementById("brkExch").value=c.exchange==="BSE"?"BSE":"NSE";
  BRK_ROWS_.forEach(function(r){
    document.getElementById("brkD_"+r[0]).value=(c.delivery||{})[r[0]];
    document.getElementById("brkI_"+r[0]).value=(c.intraday||{})[r[0]];
  });
  var mp=c.marginPrivilege||{};
  document.getElementById("brkLevOn").checked=!!mp.on;
  document.getElementById("brkLev").value=mp.intradayLeverage||5;
  updBrkLev_();
}
function readBrokerageForm_(){
  var seg=function(p){var o={};BRK_ROWS_.forEach(function(r){o[r[0]]=Number(document.getElementById(p+r[0]).value)||0;});return o;};
  return{broker:document.getElementById("brkName").value.trim(),applyToPnl:document.getElementById("brkApply").checked,exchange:document.getElementById("brkExch").value,
    delivery:seg("brkD_"),intraday:seg("brkI_"),
    marginPrivilege:{on:document.getElementById("brkLevOn").checked,intradayLeverage:Number(document.getElementById("brkLev").value)||1}};
}
function updBrkLev_(){
  var lv=Number(document.getElementById("brkLev").value)||1;
  document.getElementById("brkLevLbl").textContent=lv+"x";
  document.getElementById("brkLevHint").innerHTML='With '+lv+'x: a &#8377;1,00,000 intraday order needs only <b>&#8377;'+fmtN_(100000/lv)+'</b> of margin ('+(100/lv).toFixed(lv===Math.round(lv)?0:1)+'%), so &#8377;'+fmtN_(100000-100000/lv)+' of your margin stays free for other trades. Applies to intraday only; delivery always needs the full amount. Orders and the AI\'s trade sizing follow the same rule.';
}
function resetBrokerageDefaults(){
  if(!confirm("Replace the charges on screen with the standard Zerodha equity rates? (Nothing is saved until you press Save Brokerage.)"))return;
  var keep=readBrokerageForm_().marginPrivilege;
  var d=JSON.parse(JSON.stringify(BRK_DEFAULTS_));d.marginPrivilege=keep;
  fillBrokerageForm_(d);
}
function saveBrokerage(){
  var cfg=readBrokerageForm_();
  if(!cfg.broker){setErr_("Enter the broker name.");return;}
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){_brk=r.brokerage;_netMemo={};fillBrokerageForm_(r.brokerage);setMsg_("Brokerage settings saved \u2014 P&L and margin now use them.");loadPT(true);}
    else setErr_((r&&r.error)||"Could not save brokerage settings.");
  }).withFailureHandler(function(){setErr_("Could not reach the server.");}).saveBrokerageSettingsFromWeb(_authToken,cfg);
}

/* ── Broker (Live): your own Kite app, static-IP relay, optional automatic login ── */
var _bkState=null;
function openBrokerSettings(){
  openProfile();
  var d=document.getElementById("brokerSec");
  if(d){d.open=true;setTimeout(function(){d.scrollIntoView({block:"start"});},250);}
}
function loadBrokerSettings_(){
  google.script.run.withSuccessHandler(function(r){
    if(!r||!r.success)return;
    _bkState=r;
    var el=document.getElementById("brokerStatusLine");
    var parts=[];
    if(r.hasOwn)parts.push("Your own Kite app: <b>"+eh(r.apiKeyMasked)+"</b>");
    else if(r.usingDeploymentConnection)parts.push("Using the connection set up in this dashboard's Script Properties (<b>"+eh(r.apiKeyMasked)+"</b>) \u2014 you don't need to enter the key, secret or relay here");
    else parts.push("No broker saved yet");
    if(r.hasOwn||r.usingDeploymentConnection)parts.push(r.connectedToday?'<span style="color:var(--gr)">connected today</span>':"not connected today");
    if(r.hasProxy)parts.push("own proxy <b>saved</b>");
    if(r.autoConnect)parts.push("auto-connect <b>on</b>"+(r.lastAutoStatus?" ("+eh(r.lastAutoStatus)+")":""));
    el.innerHTML=parts.join(" &middot; ");
    document.getElementById("bkUserId").value=r.kiteUserId||"";
    document.getElementById("bkRelayUrl").value=r.relayUrl||"";
    document.getElementById("bkAuto").checked=!!r.autoConnect;
    var dep=!!r.usingDeploymentConnection;
    document.getElementById("bkApiKey").placeholder=r.hasOwn?"saved \u2014 enter a new key to replace":(dep?"set in Script Properties \u2014 leave blank":"Kite Connect API key");
    document.getElementById("bkApiSecret").placeholder=(r.hasApiSecret||dep)?(dep?"set in Script Properties \u2014 leave blank":"saved \u2014 leave blank to keep"):"API secret";
    document.getElementById("bkRelaySecret").placeholder=(r.hasRelaySecret||dep)?(dep?"set in Script Properties \u2014 leave blank":"saved \u2014 leave blank to keep"):"Shared secret your relay expects";
    document.getElementById("bkProxy").placeholder=r.hasProxy?"saved \u2014 leave blank to keep":"http://USER:PASS@host:port \u2014 leave blank to use the relay's own static IP";
    document.getElementById("bkPwd").placeholder=r.hasPassword?"saved \u2014 leave blank to keep":"Zerodha password";
    document.getElementById("bkTotp").placeholder=r.hasTotp?"saved \u2014 leave blank to keep":"TOTP secret key";
  }).withFailureHandler(function(){}).getBrokerSettingsFromWeb(_authToken);
}
function testBrokerProxy(){
  var m=document.getElementById("bkProxyMsg");m.textContent="Checking\u2026";
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    m.innerHTML=(r&&r.success)?('Your orders will leave from IP <b style="color:var(--tx)">'+eh(r.ip)+'</b> ('+(r.usingOwn?"your own proxy":"the relay's own static IP")+'). This must be the IP whitelisted in your Zerodha developer account.'):('<span style="color:var(--rd)">'+eh((r&&r.error)||"Could not check the proxy.")+'</span>');
  }).withFailureHandler(function(){m.textContent="Could not reach the server.";}).testBrokerProxyFromWeb(_authToken);
}
function clearBrokerProxy(){
  if(!confirm("Remove your own proxy? Your orders will use the relay's own static IP again."))return;
  saveBrokerSettings({clearProxy:true,apiKey:"",apiSecret:"",relaySecret:"",loginPassword:"",totpSecret:"",proxyUrl:""});
}
function askMfaIfNeeded_(){
  if(!(_bkState&&_bkState.mfaEnabled))return"";
  var c=prompt("Two-step verification: enter the 6-digit code from your authenticator app to continue.");
  return c===null?null:c;
}
function saveBrokerSettings(extra){
  var v=function(id){return document.getElementById(id).value;};
  var p={apiKey:v("bkApiKey"),apiSecret:v("bkApiSecret"),kiteUserId:v("bkUserId"),relayUrl:v("bkRelayUrl"),relaySecret:v("bkRelaySecret"),
         loginPassword:v("bkPwd"),totpSecret:v("bkTotp"),proxyUrl:v("bkProxy"),autoConnect:document.getElementById("bkAuto").checked,acknowledged:document.getElementById("bkAck").checked};
  if(extra)for(var k in extra)p[k]=extra[k];
  var code=askMfaIfNeeded_();if(code===null)return;
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){["bkApiKey","bkApiSecret","bkRelaySecret","bkProxy","bkPwd","bkTotp"].forEach(function(id){document.getElementById(id).value="";});document.getElementById("bkAck").checked=false;setMsg_("Broker settings saved.");loadBrokerSettings_();_autoConnTried=false;}
    else setErr_((r&&r.error)||"Could not save broker settings.");
  }).withFailureHandler(function(){setErr_("Could not reach the server.");}).saveBrokerSettingsFromWeb(_authToken,p,code);
}
function clearBrokerAuto(){
  if(!confirm("Delete the saved Zerodha password and TOTP secret and switch automatic login off?"))return;
  document.getElementById("bkAuto").checked=false;
  saveBrokerSettings({clearAutoCreds:true,autoConnect:false,apiKey:"",apiSecret:"",relaySecret:"",loginPassword:"",totpSecret:""});
}
function brokerConnectNow(){
  setMsg_("Connecting to Zerodha\u2026");
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){setMsg_("Connected to Zerodha for today.");loadBrokerSettings_();if(_tradingMode==="live")checkLiveConnection();}
    else setErr_((r&&r.error)||"Automatic login failed.");
  }).withFailureHandler(function(){setErr_("Could not reach the server.");}).brokerAutoConnectNowFromWeb(_authToken);
}
function removeBroker(){
  if(!confirm("Remove your saved broker details (API key, secret, relay, login)? You'll need to enter them again to trade Live."))return;
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.success){setMsg_("Broker details removed.");loadBrokerSettings_();}else setErr_((r&&r.error)||"Could not remove broker details.");
  }).withFailureHandler(function(){setErr_("Could not reach the server.");}).removeBrokerSettingsFromWeb(_authToken);
}

/* ── Two-step verification (authenticator app) ── */
function mfaShow_(which){["mfaOffBox","mfaSetupBox","mfaCodesBox","mfaOnBox"].forEach(function(id){document.getElementById(id).style.display=(id===which)?"":"none";});}
function mfaRender_(on){
  document.getElementById("mfaStatusLine").innerHTML=on?'<span style="color:var(--gr)">&#10003; On</span> \u2014 a code from your authenticator app is required at every login.':'Off \u2014 anyone with your password can log in.';
  mfaShow_(on?"mfaOnBox":"mfaOffBox");
}
function mfaStart(){
  var pw=document.getElementById("mfaPwdOn").value;
  if(!pw){setErr_("Enter your current password to start.");return;}
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(!r||!r.success){setErr_((r&&r.error)||"Could not start setup.");return;}
    document.getElementById("mfaPwdOn").value="";
    document.getElementById("mfaSecretTxt").value=r.secret; // plain key, no spaces: pastes cleanly into any app's manual entry
    var q=document.getElementById("mfaQr");q.innerHTML="";
    // v3: Microsoft Authenticator's camera needs a larger, lower-density code than Google's. Error-correction level L (fewest
    // modules) at 232 px gives roughly 6 px per module, on a white quiet zone.
    if(typeof QRCode!=="undefined"){try{new QRCode(q,{text:r.otpauthUrl,width:232,height:232,colorDark:"#000000",colorLight:"#ffffff",correctLevel:QRCode.CorrectLevel.L});}catch(e){}}
    document.getElementById("mfaConfirmCode").value="";
    mfaShow_("mfaSetupBox");
  }).withFailureHandler(function(){setErr_("Could not reach the server.");}).mfaBeginSetupFromWeb(_authToken,pw);
}
function mfaCopyKey(btn){
  var k=document.getElementById("mfaSecretTxt").value;
  try{navigator.clipboard.writeText(k).then(function(){flashBtn_(btn,"Copied \u2713");});}catch(e){window.prompt("Copy this key:",k);}
}
function flashBtn_(btn,txt){if(!btn)return;var o=btn.getAttribute("data-o")||btn.textContent;btn.setAttribute("data-o",o);btn.textContent=txt;setTimeout(function(){btn.textContent=o;},1600);}
function mfaConfirm(){
  var c=document.getElementById("mfaConfirmCode").value.trim();
  if(!c){setErr_("Enter the 6-digit code shown in your authenticator app.");return;}
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(!r||!r.success){setErr_((r&&r.error)||"That code did not work.");return;}
    document.getElementById("mfaCodesPre").textContent=r.recoveryCodes.join("\n");
    mfaShow_("mfaCodesBox");
    document.getElementById("mfaStatusLine").innerHTML='<span style="color:var(--gr)">&#10003; On</span>';
    setMsg_("Two-step verification is on. Save your recovery codes now.");
    if(_bkState)_bkState.mfaEnabled=true;
  }).withFailureHandler(function(){setErr_("Could not reach the server.");}).mfaConfirmSetupFromWeb(_authToken,c);
}
function mfaCopyCodes(){
  var t=document.getElementById("mfaCodesPre").textContent;
  try{navigator.clipboard.writeText(t).then(function(){setMsg_("Recovery codes copied.");});}catch(e){window.prompt("Copy your recovery codes:",t.replace(/\n/g," "));}
}
function mfaCodesDone(){document.getElementById("mfaCodesPre").textContent="";mfaRender_(true);}
function mfaTurnOff(){
  var pw=document.getElementById("mfaPwdOff").value,c=document.getElementById("mfaCodeOff").value.trim();
  if(!pw||!c){setErr_("Enter your password and a current authenticator (or recovery) code.");return;}
  if(!confirm("Turn off two-step verification?"))return;
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){document.getElementById("mfaPwdOff").value="";document.getElementById("mfaCodeOff").value="";mfaRender_(false);setMsg_("Two-step verification is off.");if(_bkState)_bkState.mfaEnabled=false;}
    else setErr_((r&&r.error)||"Could not turn it off.");
  }).withFailureHandler(function(){setErr_("Could not reach the server.");}).mfaDisableFromWeb(_authToken,pw,c);
}
function mfaNewCodes(){
  var pw=document.getElementById("mfaPwdOff").value,c=document.getElementById("mfaCodeOff").value.trim();
  if(!pw||!c){setErr_("Enter your password and a current authenticator code to get new recovery codes.");return;}
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){document.getElementById("mfaPwdOff").value="";document.getElementById("mfaCodeOff").value="";document.getElementById("mfaCodesPre").textContent=r.recoveryCodes.join("\n");mfaShow_("mfaCodesBox");setMsg_("New recovery codes created \u2014 the old ones no longer work.");}
    else setErr_((r&&r.error)||"Could not create new codes.");
  }).withFailureHandler(function(){setErr_("Could not reach the server.");}).mfaRegenerateRecoveryFromWeb(_authToken,pw,c);
}

var _trafRange=30;
function setTrafficRange(days){
  _trafRange=days;
  document.querySelectorAll("#trafRangeBtns .tybtn").forEach(function(b){b.classList.toggle("sel",Number(b.dataset.r)===days);});
  loadTrafficStats_();
}
function openTraffic(){
  document.getElementById("trafficM").classList.add("open");
  document.querySelectorAll("#trafRangeBtns .tybtn").forEach(function(b){b.classList.toggle("sel",Number(b.dataset.r)===_trafRange);});
  loadTrafficStats_();
}
function countUp_(id,target){
  target=Number(target)||0;
  var el=document.getElementById(id);if(!el)return;
  var start=0,dur=600,t0=null;
  function step(ts){
    if(!t0)t0=ts;
    var p=Math.min(1,(ts-t0)/dur);
    var eased=1-Math.pow(1-p,3);
    el.textContent=Math.round(start+(target-start)*eased);
    if(p<1)requestAnimationFrame(step); else el.textContent=target;
  }
  requestAnimationFrame(step);
}
function loadTrafficStats_(){
  var chartEl=document.getElementById("trafChart");
  chartEl.innerHTML='<div style="color:var(--dm);font-size:11px;padding:20px 0;text-align:center">Loading...</div>';
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(!r||!r.success){chartEl.innerHTML='<div style="color:var(--dm);font-size:11px">'+eh((r&&r.error)||"Not authorized.")+'</div>';return;}
    countUp_("trafTotal",r.totalVisits);
    countUp_("trafUnique",r.uniqueUsers);
    countUp_("trafToday",r.today);
    countUp_("trafRange",r.rangeTotal);
    document.getElementById("trafRangeLbl").textContent=(_trafRange?_trafRange+"D Total":"All-time Total")+(r.firstDate?" (since "+r.firstDate+")":"");
    if(!r.days.length){chartEl.innerHTML='<div style="color:var(--dm);font-size:11px;padding:20px 0;text-align:center">No visits logged yet.</div>';return;}
    renderTrafficChart_(chartEl,r.days,r.dayCounts);
  }).withFailureHandler(function(){chartEl.innerHTML='<div style="color:var(--dm);font-size:11px">Server error.</div>';}).getTrafficStatsFromWeb(_authToken,_trafRange);
}
function renderTrafficChart_(el,days,counts){
  var W=520,H=200,padL=32,padB=28,padT=10,padR=10;
  var plotW=W-padL-padR,plotH=H-padT-padB;
  var max=Math.max.apply(null,counts.concat([1]));
  var n=days.length;
  var bw=Math.max(2,plotW/n*0.7);
  var step=plotW/n;
  var gridLines="",yLabels="";
  var ySteps=4;
  for(var g=0;g<=ySteps;g++){
    var val=Math.round(max*g/ySteps);
    var y=padT+plotH-(plotH*g/ySteps);
    gridLines+='<line x1="'+padL+'" y1="'+y+'" x2="'+(W-padR)+'" y2="'+y+'" stroke="var(--bd)" stroke-width="1" opacity="0.5"/>';
    yLabels+='<text x="'+(padL-6)+'" y="'+(y+3)+'" font-size="9" fill="var(--dm)" text-anchor="end">'+val+'</text>';
  }
  var labelEvery=Math.ceil(n/8);
  var bars="",xLabels="",baseline=padT+plotH;
  days.forEach(function(d,i){
    var x=padL+i*step+(step-bw)/2;
    var h=(counts[i]/max)*plotH;
    bars+='<rect data-fy="'+(baseline-Math.max(h,1)).toFixed(1)+'" data-fh="'+Math.max(h,1).toFixed(1)+'" x="'+x.toFixed(1)+'" y="'+baseline+'" width="'+bw.toFixed(1)+'" height="0" fill="var(--ac)" rx="2" class="trafBar" style="transition:height .5s cubic-bezier(.2,.7,.3,1) '+(i*15)+'ms,y .5s cubic-bezier(.2,.7,.3,1) '+(i*15)+'ms,opacity .15s"><title>'+d+': '+counts[i]+' visit'+(counts[i]===1?'':'s')+'</title></rect>';
    if(i%labelEvery===0||i===n-1){
      xLabels+='<text x="'+(x+bw/2).toFixed(1)+'" y="'+(H-padB+12)+'" font-size="9" fill="var(--dm)" text-anchor="middle">'+d.slice(5)+'</text>';
    }
  });
  el.innerHTML='<svg viewBox="0 0 '+W+' '+H+'" style="width:100%;height:auto;display:block;opacity:0;transition:opacity .3s" preserveAspectRatio="xMidYMid meet">'+gridLines+yLabels+bars+xLabels+'</svg>'+
    '<style>.trafBar:hover{opacity:0.7}</style>';
  var svg=el.querySelector("svg");
  requestAnimationFrame(function(){
    svg.style.opacity=1;
    svg.querySelectorAll(".trafBar").forEach(function(r){r.setAttribute("y",r.getAttribute("data-fy"));r.setAttribute("height",r.getAttribute("data-fh"));});
  });
}
function closeTraffic(){document.getElementById("trafficM").classList.remove("open");}
function saveProfileName(){
  var name=document.getElementById("profName").value.trim();
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){showErr("profMsg","Name updated.");document.getElementById("profMsg").classList.add("show");document.getElementById("profErr").classList.remove("show");}
    else{showErr("profErr",(r&&r.error)||"Could not update name.");}
  }).withFailureHandler(function(){}).updateProfileFromWeb(_authToken,name);
}
function saveProfilePassword(){
  var cur=document.getElementById("profCurPass").value,p1=document.getElementById("profNewPass").value,p2=document.getElementById("profNewPass2").value;
  if(!cur||!p1){showErr("profErr","Fill in all password fields.");return;}
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.sessionExpired){forceReauth();return;}
    if(r&&r.success){document.getElementById("profCurPass").value="";document.getElementById("profNewPass").value="";document.getElementById("profNewPass2").value="";showErr("profMsg","Password changed.");document.getElementById("profMsg").classList.add("show");document.getElementById("profErr").classList.remove("show");}
    else{showErr("profErr",(r&&r.error)||"Could not change password.");}
  }).withFailureHandler(function(){}).changePasswordFromWeb(_authToken,cur,p1,p2);
}
function confirmDeleteAccount(){
  var pass=prompt("This permanently deletes your account, portfolio, and watchlist. Enter your password to confirm:");
  if(!pass)return;
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.success){localStorage.removeItem("sdToken");location.reload();}
    else{alert((r&&r.error)||"Could not delete account.");}
  }).withFailureHandler(function(){}).deleteAccountFromWeb(_authToken,pass);
}
function forceReauth(){localStorage.removeItem("sdToken");_authToken=null;try{sessionStorage.setItem("sdLogoutReason","Your session ended — this may be because your account was used to log in elsewhere (only one active session is allowed per account).");}catch(e){}location.reload();}

/* ══ INIT — gated behind auth ══
   initDashboardApp() below is byte-for-byte the same logic that used to
   run directly on window load. It now only runs once a session is
   confirmed (either an existing valid token, or a fresh login/signup) —
   nothing about what it DOES has changed, only when it fires. */
function initDashboardApp(){
  var td=todayStr();_gDate=td;document.getElementById("gDate").value=td;
  restoreTabCfg("home");
  switchTopTab("home");
  preloadAll();
  setTimeout(fetchNews,3500);
  setTimeout(function(){ensureNseMaster(function(){});},1500);
  setInterval(function(){
    refreshNseMaster_(function(){
      if(_wlRows&&_wlRows.length&&document.getElementById("cnt-watchlist").style.display!=="none"){_wlRows=mergeNseData_(_wlRows);renderWL();}
      if(_ptRows&&_ptRows.length&&document.getElementById("cnt-portfolio").style.display!=="none"){_ptRows=mergeNseData_(_ptRows);renderPT();}
    });
  },30000);
  // Pre-load watchlist symbol set so star icons are highlighted correctly
  // across all tabs without requiring the user to open the Watchlist tab first
  setTimeout(function(){
    google.script.run.withSuccessHandler(function(r){
      if(r&&r.success&&r.rows){_wlRows=r.rows;_wlSymSet.clear();r.rows.forEach(function(row){_wlSymSet.add(row.symbol);});}
    }).withFailureHandler(function(){}).getWatchlistFromWeb(_authToken);
  },2000);
  setTimeout(function(){
    google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(){}).logVisitFromWeb(_authToken);
    google.script.run.withSuccessHandler(function(r){
      if(r&&r.isAdmin)document.getElementById("trafficBtn").style.display="";
    }).withFailureHandler(function(){}).isAdminFromWeb(_authToken);
  },2500);
  schedRef();
  attachAutocomplete("wlSym",function(item){document.getElementById("wlNote")&&document.getElementById("wlNote").focus();});
  attachAutocomplete("ptSym",function(item){if(item.ltp)document.getElementById("ptPrice").placeholder="LTP: "+item.ltp;document.getElementById("ptQty").focus();});
  attachAutocomplete("liveOrdSymbol",function(item){
    _liveOrdLtp=item.ltp||null;
    document.getElementById("liveOrdPrice").placeholder=_liveOrdLtp?("LTP: "+_liveOrdLtp):"Price";
    document.getElementById("liveOrdQty").focus();
  });
  loadReferralInfo();
  loadNtdFlag_(); // v3
  ["obQty","obPrice"].forEach(function(id){var el=document.getElementById(id);if(el)el.addEventListener("input",updOrderEstimate);});
  // Top-edge Order Book panel — same hover-delay pattern as the right-edge AI panel
  var obEdge=document.getElementById("obEdge"),_obEdgeTmr=null;
  if(obEdge){
    obEdge.addEventListener("mouseenter",function(){_obEdgeTmr=setTimeout(openOB,600);});
    obEdge.addEventListener("mouseleave",function(){clearTimeout(_obEdgeTmr);});
    window.addEventListener("wheel",function(){clearTimeout(_obEdgeTmr);},{passive:true});
  }
}
var _refLink="";
function loadReferralInfo(){
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.success){
      _myReferralCode=r.code;
      var base=(r.appUrl&&r.appUrl.indexOf("http")===0)?r.appUrl:(window.location.origin+window.location.pathname);
      _refLink=base+(base.indexOf("?")>=0?"&":"?")+"ref="+r.code;
      var fl=document.getElementById("refLinkFull");if(fl)fl.value=_refLink;
      var fl2=document.getElementById("refLinkShare");if(fl2)fl2.value=_refLink;
      var rc=document.getElementById("refCount");if(rc)rc.textContent="("+r.count+" joined via your link)";
      var rc2=document.getElementById("refCountSh");if(rc2)rc2.textContent="("+r.count+" joined)";
    }
  }).withFailureHandler(function(){}).getReferralInfoFromWeb(_authToken);
}
function copyReferralLink(){
  var el=document.getElementById("refLinkShare").value?document.getElementById("refLinkShare"):document.getElementById("refLinkFull");
  if(!el||!el.value)return;
  el.select();el.setSelectionRange(0,99999);
  try{document.execCommand("copy");}catch(e){}
  var rt=document.getElementById("refTxt");
  if(rt){var orig=rt.textContent;rt.textContent="Copied!";setTimeout(function(){rt.textContent=orig;},1500);}
}
function openShare(){document.getElementById("shareM").classList.add("open");if(!_refLink)loadReferralInfo();}
function closeShare(){document.getElementById("shareM").classList.remove("open");}
function shareTo(platform){
  if(!_refLink){loadReferralInfo();setTimeout(function(){shareTo(platform);},400);return;}
  var msg="Check out this NSE/BSE stock signal dashboard — track intraday, short-term & long-term signals with AI insights.";
  var link=encodeURIComponent(_refLink), text=encodeURIComponent(msg);
  var urls={
    whatsapp:"https://wa.me/?text="+text+"%20"+link,
    twitter:"https://twitter.com/intent/tweet?text="+text+"&url="+link,
    facebook:"https://www.facebook.com/sharer/sharer.php?u="+link,
    linkedin:"https://www.linkedin.com/sharing/share-offsite/?url="+link,
    telegram:"https://t.me/share/url?url="+link+"&text="+text,
    email:"mailto:?subject="+encodeURIComponent("Check this out")+"&body="+text+"%20"+link
  };
  if(urls[platform])window.open(urls[platform],"_blank");
}
window.addEventListener("load",function(){
  var token=localStorage.getItem("sdToken");
  if(!token){showAuthGate();return;}
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.valid){_authToken=token;_tradingAccountPref=r.tradingAccountPref||"mock";hideAuthGate();initDashboardApp();}
    else{localStorage.removeItem("sdToken");showAuthGate();}
  }).withFailureHandler(function(){showAuthGate();}).validateSessionFromWeb(token);
});

</script>
</body>
</html>

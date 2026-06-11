import re

with open('website/index.html', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update cursor html
content = re.sub(
    r'<div id="dynamic-cursor" class="absolute z-40 pointer-events-none mix-blend-screen transition-transform duration-300">',
    r'<div id="dynamic-cursor" class="absolute z-40 pointer-events-none mix-blend-screen">',
    content
)

# 2. Update glow html
content = re.sub(
    r'<div id="ambient-glow" class="fixed w-\[800px\] h-\[800px\] rounded-full bg-blue-500/10 blur-\[150px\] pointer-events-none z-0 left-1/2 top-1/2 -translate-x-1/2 -translate-y-1/2 transition-all duration-1000 opacity-[0-9]+"></div>',
    r'<div id="ambient-glow" class="fixed w-[800px] h-[800px] rounded-full bg-blue-500/10 blur-[150px] pointer-events-none z-0 left-1/2 top-1/2 -translate-x-1/2 -translate-y-1/2 opacity-80"></div>',
    content
)

# 3. Add pointer-events-none to layers initially to avoid layout block before they animate in
content = re.sub(
    r'id="layer-word-sheet" class="layer-3d w-\[440px\] h-\[600px\] bg-white rounded-xl shadow-2xl overflow-hidden opacity-0"',
    r'id="layer-word-sheet" class="layer-3d w-[440px] h-[600px] bg-white rounded-xl shadow-2xl overflow-hidden opacity-0 pointer-events-none"',
    content
)
content = re.sub(
    r'id="layer-word-chat" class="layer-3d w-\[350px\] glass-panel border border-blue-500/20 rounded-2xl p-5 shadow-2xl opacity-0"',
    r'id="layer-word-chat" class="layer-3d w-[350px] glass-panel border border-blue-500/20 rounded-2xl p-5 shadow-2xl opacity-0 pointer-events-none"',
    content
)
content = re.sub(
    r'id="layer-word-revision" class="layer-3d w-\[280px\] glass-panel border border-emerald-500/30 rounded-xl p-4 shadow-2xl opacity-0"',
    r'id="layer-word-revision" class="layer-3d w-[280px] glass-panel border border-emerald-500/30 rounded-xl p-4 shadow-2xl opacity-0 pointer-events-none"',
    content
)
content = re.sub(
    r'id="layer-excel" class="layer-3d w-\[640px\] glass-panel border border-emerald-500/20 rounded-2xl p-5 shadow-2xl opacity-0"',
    r'id="layer-excel" class="layer-3d w-[640px] glass-panel border border-emerald-500/20 rounded-2xl p-5 shadow-2xl opacity-0 pointer-events-none"',
    content
)
content = re.sub(
    r'id="layer-ppt" class="layer-3d w-full max-w-5xl opacity-0 flex flex-col items-center justify-center"',
    r'id="layer-ppt" class="layer-3d w-full max-w-5xl opacity-0 flex flex-col items-center justify-center pointer-events-none"',
    content
)
content = re.sub(
    r'id="layer-desktop" class="layer-3d w-\[900px\] h-\[550px\] glass-panel border border-white/10 rounded-2xl shadow-\[0_50px_100px_rgba\(0,0,0,0\.8\)\] opacity-0 flex flex-col overflow-hidden"',
    r'id="layer-desktop" class="layer-3d w-[900px] h-[550px] glass-panel border border-white/10 rounded-2xl shadow-[0_50px_100px_rgba(0,0,0,0.8)] opacity-0 flex flex-col overflow-hidden pointer-events-none"',
    content
)

# 4. Replace cursorState definition
old_cursor_state = """    let cursorState = {
      shape: 'dot',     // dot -> quill -> prism -> projector -> portal
      scale: 1.0,
      glowColor: 'rgba(59, 130, 246, 0.4)', // Default blue
      x: 0,
      y: 0,
      rotation: 0
    };"""

new_cursor_state = """    let cursorState = {
      radius: 50,
      width: 0,
      height: 0,
      rectAlpha: 0,
      rayLength: 0,
      rayAlpha: 0,
      innerRadius: 0,
      orbitRadius: 0,
      orbitAlpha: 0,
      colorR: 59,
      colorG: 130,
      colorB: 246,
      colorA: 0.8,
      scale: 1.0,
      x: 0,
      y: 0,
      rotation: 0
    };"""
content = content.replace(old_cursor_state, new_cursor_state)


# 5. Replace drawMorphingCursor definition
old_draw_morphing = """    function drawMorphingCursor() {
      cctx.clearRect(0, 0, 400, 400);
      const cx = 200;
      const cy = 200;

      cctx.save();
      cctx.translate(cx, cy);
      cctx.rotate(cursorState.rotation);

      if (cursorState.shape === 'dot') {
        // Glowing celestial spark dot (Hero & Portal start)
        const pulse = 1.0 + Math.sin(Date.now() / 250) * 0.08;
        const grad = cctx.createRadialGradient(0, 0, 1, 0, 0, 50 * cursorState.scale * pulse);
        grad.addColorStop(0, 'rgba(255, 255, 255, 1)');
        grad.addColorStop(0.2, 'rgba(59, 130, 246, 0.8)');
        grad.addColorStop(0.5, 'rgba(59, 130, 246, 0.25)');
        grad.addColorStop(1, 'rgba(59, 130, 246, 0)');

        cctx.beginPath();
        cctx.arc(0, 0, 50 * cursorState.scale * pulse, 0, Math.PI * 2);
        cctx.fillStyle = grad;
        cctx.fill();
      }
      else if (cursorState.shape === 'quill') {
        // Blinking typing vertical beam (Word Document stage)
        const blinkAlpha = Math.abs(Math.sin(Date.now() / 280));
        cctx.beginPath();
        cctx.rect(-3, -70 * cursorState.scale, 6, 140 * cursorState.scale);
        cctx.fillStyle = `rgba(59, 130, 246, ${0.2 + blinkAlpha * 0.75})`;
        cctx.shadowColor = 'rgba(59, 130, 246, 0.8)';
        cctx.shadowBlur = 18;
        cctx.fill();
      }
      else if (cursorState.shape === 'prism') {
        // Horizontal emerald grid scanner (Excel Grid stage)
        const sweepY = Math.sin(Date.now() / 350) * 60;

        // Scan line
        cctx.beginPath();
        cctx.rect(-150, sweepY - 1.5, 300, 3);
        cctx.fillStyle = 'rgba(16, 185, 129, 0.9)';
        cctx.shadowColor = 'rgba(16, 185, 129, 0.9)';
        cctx.shadowBlur = 15;
        cctx.fill();

        // Tilted preview bounds
        cctx.shadowBlur = 0;
        cctx.strokeStyle = 'rgba(16, 185, 129, 0.25)';
        cctx.lineWidth = 1;
        cctx.strokeRect(-120, -70, 240, 140);
      }
      else if (cursorState.shape === 'projector') {
        // Multi-dimensional projecting vectors (PPT generation stage)
        const pulse = 1.0 + Math.sin(Date.now() / 200) * 0.12;

        // Dynamic geometric lines projecting outwards
        cctx.strokeStyle = 'rgba(139, 92, 246, 0.4)';
        cctx.lineWidth = 1.5;
        cctx.shadowColor = 'rgba(139, 92, 246, 0.6)';
        cctx.shadowBlur = 10;

        cctx.beginPath();
        for (let i = 0; i < 4; i++) {
          const angle = (i * Math.PI) / 2 + (Date.now() / 2000);
          const rx = Math.cos(angle) * 110 * pulse * cursorState.scale;
          const ry = Math.sin(angle) * 110 * pulse * cursorState.scale;
          cctx.moveTo(0, 0);
          cctx.lineTo(rx, ry);
        }
        cctx.stroke();

        // Inner glowing core
        cctx.beginPath();
        cctx.arc(0, 0, 14, 0, Math.PI * 2);
        cctx.fillStyle = '#fff';
        cctx.shadowColor = 'rgba(139, 92, 246, 1)';
        cctx.shadowBlur = 20;
        cctx.fill();
      }
      else if (cursorState.shape === 'portal') {
        // spinning glowing portal rings (Converging / Final CTA stage)
        const rotationSpeed = Date.now() / 800;

        // Inner logo circle
        cctx.beginPath();
        cctx.arc(0, 0, 40 * cursorState.scale, 0, Math.PI * 2);
        cctx.strokeStyle = 'rgba(59, 130, 246, 0.6)';
        cctx.lineWidth = 2;
        cctx.shadowColor = 'rgba(59, 130, 246, 0.7)';
        cctx.shadowBlur = 15;
        cctx.stroke();

        // Outer spinning orbital dots
        for (let i = 0; i < 3; i++) {
          const angle = rotationSpeed + (i * Math.PI * 2) / 3;
          const ox = Math.cos(angle) * 70 * cursorState.scale;
          const oy = Math.sin(angle) * 70 * cursorState.scale;
          cctx.beginPath();
          cctx.arc(ox, oy, 6, 0, Math.PI * 2);
          cctx.fillStyle = '#60a5fa';
          cctx.fill();
        }
      }

      cctx.restore();
      requestAnimationFrame(drawMorphingCursor);
    }"""

new_draw_morphing = """    function drawMorphingCursor() {
      cctx.clearRect(0, 0, 400, 400);
      const cx = 200;
      const cy = 200;

      cctx.save();
      cctx.translate(cx, cy);
      cctx.rotate(cursorState.rotation);

      // 1. Draw glowing celestial spark dot
      if (cursorState.radius > 0.1 && cursorState.colorA > 0.01) {
        const pulse = 1.0 + Math.sin(Date.now() / 250) * 0.08;
        const r = cursorState.radius * cursorState.scale * pulse;
        const grad = cctx.createRadialGradient(0, 0, 1, 0, 0, Math.max(1, r));
        grad.addColorStop(0, `rgba(255, 255, 255, ${cursorState.colorA})`);
        grad.addColorStop(0.2, `rgba(${Math.round(cursorState.colorR)}, ${Math.round(cursorState.colorG)}, ${Math.round(cursorState.colorB)}, ${cursorState.colorA})`);
        grad.addColorStop(0.5, `rgba(${Math.round(cursorState.colorR)}, ${Math.round(cursorState.colorG)}, ${Math.round(cursorState.colorB)}, ${cursorState.colorA * 0.3})`);
        grad.addColorStop(1, `rgba(${Math.round(cursorState.colorR)}, ${Math.round(cursorState.colorG)}, ${Math.round(cursorState.colorB)}, 0)`);

        cctx.beginPath();
        cctx.arc(0, 0, Math.max(0, r), 0, Math.PI * 2);
        cctx.fillStyle = grad;
        cctx.fill();
      }

      // 2. Draw rectangular elements (quill beam or prism scan line)
      if (cursorState.rectAlpha > 0.01) {
        cctx.save();
        const blinkAlpha = Math.abs(Math.sin(Date.now() / 280));
        // quill blinks, prism scan line doesn't. We can use height > width to detect quill
        const factor = cursorState.height > cursorState.width ? (0.2 + blinkAlpha * 0.75) : 1.0;

        cctx.beginPath();
        cctx.rect(-cursorState.width / 2 * cursorState.scale, -cursorState.height / 2 * cursorState.scale, cursorState.width * cursorState.scale, cursorState.height * cursorState.scale);
        cctx.fillStyle = `rgba(${Math.round(cursorState.colorR)}, ${Math.round(cursorState.colorG)}, ${Math.round(cursorState.colorB)}, ${cursorState.rectAlpha * factor})`;
        cctx.shadowColor = `rgba(${Math.round(cursorState.colorR)}, ${Math.round(cursorState.colorG)}, ${Math.round(cursorState.colorB)}, 0.8)`;
        cctx.shadowBlur = 15;
        cctx.fill();

        // If it's the prism scanner (width > 150), also draw the bounds and sweep scan line
        if (cursorState.width > 150) {
          const boundsW = cursorState.width * 0.8;
          cctx.shadowBlur = 0;
          cctx.strokeStyle = `rgba(${Math.round(cursorState.colorR)}, ${Math.round(cursorState.colorG)}, ${Math.round(cursorState.colorB)}, ${cursorState.rectAlpha * 0.25})`;
          cctx.lineWidth = 1;
          cctx.strokeRect(-boundsW / 2 * cursorState.scale, -70 * cursorState.scale, boundsW * cursorState.scale, 140 * cursorState.scale);

          // Sweeping scan line
          const sweepY = Math.sin(Date.now() / 350) * 60;
          cctx.beginPath();
          cctx.rect(-boundsW / 2 * cursorState.scale, sweepY - 1.5, boundsW * cursorState.scale, 3);
          cctx.fillStyle = `rgba(${Math.round(cursorState.colorR)}, ${Math.round(cursorState.colorG)}, ${Math.round(cursorState.colorB)}, ${cursorState.rectAlpha * 0.9})`;
          cctx.shadowColor = `rgba(${Math.round(cursorState.colorR)}, ${Math.round(cursorState.colorG)}, ${Math.round(cursorState.colorB)}, 0.9)`;
          cctx.shadowBlur = 15;
          cctx.fill();
        }
        cctx.restore();
      }

      // 3. Draw projecting vectors (rays)
      if (cursorState.rayAlpha > 0.01) {
        cctx.save();
        const pulse = 1.0 + Math.sin(Date.now() / 200) * 0.12;
        cctx.strokeStyle = `rgba(${Math.round(cursorState.colorR)}, ${Math.round(cursorState.colorG)}, ${Math.round(cursorState.colorB)}, ${cursorState.rayAlpha * 0.4})`;
        cctx.lineWidth = 1.5;
        cctx.shadowColor = `rgba(${Math.round(cursorState.colorR)}, ${Math.round(cursorState.colorG)}, ${Math.round(cursorState.colorB)}, ${cursorState.rayAlpha * 0.6})`;
        cctx.shadowBlur = 10;

        cctx.beginPath();
        for (let i = 0; i < 4; i++) {
          const angle = (i * Math.PI) / 2 + (Date.now() / 2000);
          const rx = Math.cos(angle) * cursorState.rayLength * pulse * cursorState.scale;
          const ry = Math.sin(angle) * cursorState.rayLength * pulse * cursorState.scale;
          cctx.moveTo(0, 0);
          cctx.lineTo(rx, ry);
        }
        cctx.stroke();

        // Inner glowing core
        cctx.beginPath();
        cctx.arc(0, 0, 14 * cursorState.scale, 0, Math.PI * 2);
        cctx.fillStyle = `rgba(255, 255, 255, ${cursorState.rayAlpha})`;
        cctx.shadowColor = `rgba(${Math.round(cursorState.colorR)}, ${Math.round(cursorState.colorG)}, ${Math.round(cursorState.colorB)}, ${cursorState.rayAlpha})`;
        cctx.shadowBlur = 20;
        cctx.fill();
        cctx.restore();
      }

      // 4. Draw spinning portal rings
      if (cursorState.orbitAlpha > 0.01) {
        cctx.save();
        const rotationSpeed = Date.now() / 800;

        // Inner circle
        cctx.beginPath();
        cctx.arc(0, 0, cursorState.innerRadius * cursorState.scale, 0, Math.PI * 2);
        cctx.strokeStyle = `rgba(${Math.round(cursorState.colorR)}, ${Math.round(cursorState.colorG)}, ${Math.round(cursorState.colorB)}, ${cursorState.orbitAlpha * 0.6})`;
        cctx.lineWidth = 2;
        cctx.shadowColor = `rgba(${Math.round(cursorState.colorR)}, ${Math.round(cursorState.colorG)}, ${Math.round(cursorState.colorB)}, ${cursorState.orbitAlpha * 0.7})`;
        cctx.shadowBlur = 15;
        cctx.stroke();

        // Outer spinning orbital dots
        for (let i = 0; i < 3; i++) {
          const angle = rotationSpeed + (i * Math.PI * 2) / 3;
          const ox = Math.cos(angle) * cursorState.orbitRadius * cursorState.scale;
          const oy = Math.sin(angle) * cursorState.orbitRadius * cursorState.scale;
          cctx.beginPath();
          cctx.arc(ox, oy, 6 * cursorState.scale, 0, Math.PI * 2);
          cctx.fillStyle = `rgba(96, 165, 250, ${cursorState.orbitAlpha})`;
          cctx.fill();
        }
        cctx.restore();
      }

      cctx.restore();
      requestAnimationFrame(drawMorphingCursor);
    }"""
content = content.replace(old_draw_morphing, new_draw_morphing)

# 6. Replace GSAP Timeline completely
old_timeline_block = content[content.find("// Scroll typing tween helper function") : content.find("</script>")]

new_timeline_block = """// Scroll typing tween helper function
    function scrollTypingTween(timeline, elementId, fullText, duration, position) {
      const el = document.getElementById(elementId);
      if (!el) return;

      const textObj = { length: 0 };

      timeline.to(textObj, {
        length: fullText.length,
        duration: duration,
        ease: "none",
        onUpdate: function() {
          const charsToShow = Math.floor(textObj.length);
          el.textContent = fullText.slice(0, charsToShow);
          if (charsToShow > 0 && charsToShow < fullText.length) {
            el.classList.add('typewriter-caret');
          } else {
            el.classList.remove('typewriter-caret');
          }
        }
      }, position);
    }

    // Toggle Excel Grid text replacement
    function toggleExcelCells(repaired) {
      const err1 = document.getElementById('excel-err-1');
      const err2 = document.getElementById('excel-err-2');
      const err3 = document.getElementById('excel-err-3');
      const err4 = document.getElementById('excel-err-4');
      const sum1 = document.getElementById('excel-sum-1');
      const sum2 = document.getElementById('excel-sum-2');

      if (!err1) return;

      if (repaired) {
        err1.textContent = "¥16,940";
        err1.className = "p-2.5 bg-emerald-500/10 text-emerald-400 font-semibold text-right transition-colors duration-700";

        err2.textContent = "¥23,320";
        err2.className = "p-2.5 bg-emerald-500/10 text-emerald-400 font-semibold text-right transition-colors duration-700";

        err3.textContent = "¥12,450";
        err3.className = "p-2.5 bg-emerald-500/10 text-emerald-400 font-semibold text-right transition-colors duration-700";

        err4.textContent = "¥52,870";
        err4.className = "p-2.5 bg-emerald-500/10 text-emerald-400 font-semibold text-right transition-colors duration-700";

        sum1.textContent = "¥48,710";
        sum1.className = "p-2.5 bg-slate-900/60 text-slate-300 font-bold text-right";

        sum2.textContent = "¥48,400";
        sum2.className = "p-2.5 bg-slate-900/60 text-slate-300 font-bold text-right";
      } else {
        err1.textContent = "#REF!";
        err1.className = "p-2.5 bg-red-950/30 text-red-400 font-bold text-center border border-red-500/20 excel-error-cell transition-all duration-700";

        err2.textContent = "#VALUE!";
        err2.className = "p-2.5 bg-red-950/30 text-red-400 font-bold text-center border border-red-500/20 excel-error-cell transition-all duration-700";

        err3.textContent = "#REF!";
        err3.className = "p-2.5 bg-red-950/30 text-red-400 font-bold text-center border border-red-500/20 excel-error-cell transition-all duration-700";

        err4.textContent = "#N/A";
        err4.className = "p-2.5 bg-red-950/30 text-red-400 font-bold text-center border border-red-500/20 excel-error-cell transition-all duration-700";

        sum1.textContent = "--";
        sum1.className = "p-2.5 bg-slate-950/40 text-slate-500 text-center";

        sum2.textContent = "--";
        sum2.className = "p-2.5 bg-slate-950/40 text-slate-500 text-center";
      }
    }

    const mainTimeline = gsap.timeline({
      scrollTrigger: {
        trigger: '#scrolly-trigger',
        start: 'top top',
        end: 'bottom bottom',
        pin: '#scrolly-pinned',
        scrub: 1.2,
        invalidateOnRefresh: true
      }
    });

    // Setup sequence milestones and state updates
    mainTimeline
      // --- PHASE 1: HERO STATE ---
      .addLabel('hero', 0)
      .to('#layer-hero', { opacity: 1, z: 0, duration: 2, ease: 'power2.out' })
      .to('#layer-hero', { pointerEvents: 'auto', duration: 0.1 }, 0)
      .to(cursorState, {
        radius: 50,
        width: 0,
        height: 0,
        rectAlpha: 0,
        rayLength: 0,
        rayAlpha: 0,
        innerRadius: 0,
        orbitRadius: 0,
        orbitAlpha: 0,
        colorR: 59,
        colorG: 130,
        colorB: 246,
        colorA: 0.8,
        scale: 0.9,
        rotation: 0,
        duration: 2
      }, 0)
      .to('#dynamic-cursor', { x: 0, y: -180, duration: 2, ease: 'power2.out' }, 0)

      // Transition from Hero to Word Editor
      .to('#layer-hero', { opacity: 0, y: -100, z: -200, duration: 2, ease: 'power2.inOut' }, 2)
      .to('#layer-hero', { pointerEvents: 'none', duration: 0.1 }, 2)

      // Blend cursor morph properties smoothly to quill beam
      .to(cursorState, {
        radius: 0,
        width: 6,
        height: 140,
        rectAlpha: 0.85,
        rayLength: 0,
        rayAlpha: 0,
        innerRadius: 0,
        orbitRadius: 0,
        orbitAlpha: 0,
        colorR: 59,
        colorG: 130,
        colorB: 246,
        colorA: 0.0,
        scale: 0.75,
        duration: 2,
        ease: 'power2.inOut'
      }, 2)
      .to('#dynamic-cursor', { x: 200, y: 40, duration: 2, ease: 'power2.inOut' }, 2)

      // --- PHASE 2: WORD PROCESSOR ---
      .addLabel('word', 4)
      .to('#layer-word-sheet', { opacity: 1, x: -220, y: 0, z: 0, rotationY: 12, duration: 2, ease: 'power3.out' }, 4)
      .to('#layer-word-sheet', { pointerEvents: 'auto', duration: 0.1 }, 4)
      .to('#layer-word-chat', { opacity: 1, x: 260, y: 30, z: 40, rotationY: -10, duration: 2, ease: 'power3.out' }, 4)
      .to('#layer-word-chat', { pointerEvents: 'auto', duration: 0.1 }, 4)

      // Scroll bound typing user input
      .call(() => {
        // Safe check for back/forward scrub matching
        const responseEl = document.getElementById('word-chat-response');
        if (responseEl) {
          gsap.set(responseEl, { opacity: 0, y: 15 });
        }
      }, null, 4.5)

      .call(() => {
        // Spawn mathematical particles flying from Chat panel to Document sheet in scroll trigger
        const chatEl = document.getElementById('layer-word-chat');
        const docEl = document.getElementById('layer-word-sheet');
        if (chatEl && docEl) {
          const chatRect = chatEl.getBoundingClientRect();
          const docRect = docEl.getBoundingClientRect();
          spawnWordParticles(
            chatRect.left + 50, chatRect.top + 160,
            docRect.left + 220, docRect.top + 280,
            20
          );
        }
      }, null, 6.0)

      // User typing
      .call(() => { scrollTypingTween(mainTimeline, 'word-chat-input', '/润色 补充第一段销售业绩，融入真实Q1报表数据。', 1.8, 4.5); }, null, 4.4)

      // AI Response reveals
      .to('#word-chat-response', { opacity: 1, y: 0, duration: 0.8 }, 6.2)

      // Cursor moves to document sheet and transforms slightly
      .to('#dynamic-cursor', { x: -200, y: 40, scale: 0.8, duration: 1.5, ease: 'power2.inOut' }, 6.5)

      // Document title typing
      .call(() => { scrollTypingTween(mainTimeline, 'word-doc-title', '2026年度业务数据报告.docx', 1.8, 6.5); }, null, 6.4)

      // Diff displays
      .to('#word-diff-del', { opacity: 1, x: 0, duration: 0.8 }, 8.2)
      .to('#word-diff-add', { opacity: 1, x: 0, duration: 0.8 }, 8.5)

      // Revision Panel floats in
      .to('#layer-word-revision', { opacity: 1, x: -330, y: 190, z: 80, duration: 1.0, ease: 'back.out(1.2)' }, 8.8)
      .to('#layer-word-revision', { pointerEvents: 'auto', duration: 0.1 }, 8.8)
      .to({}, { duration: 1.5 }) // Hold frame

      // Transition from Word to Excel Spreadsheet
      .to('#layer-word-sheet', { opacity: 0, x: -400, z: -300, rotationY: 45, duration: 2, ease: 'power2.in' }, 11)
      .to('#layer-word-sheet', { pointerEvents: 'none', duration: 0.1 }, 11)
      .to('#layer-word-chat', { opacity: 0, x: 400, z: -300, rotationY: -45, duration: 2, ease: 'power2.in' }, 11)
      .to('#layer-word-chat', { pointerEvents: 'none', duration: 0.1 }, 11)
      .to('#layer-word-revision', { opacity: 0, scale: 0.6, duration: 1.5 }, 11)
      .to('#layer-word-revision', { pointerEvents: 'none', duration: 0.1 }, 11)

      // Blend cursor morph properties smoothly to prism grid scanner (emerald)
      .to(cursorState, {
        radius: 0,
        width: 300,
        height: 3,
        rectAlpha: 0.9,
        rayLength: 0,
        rayAlpha: 0,
        innerRadius: 0,
        orbitRadius: 0,
        orbitAlpha: 0,
        colorR: 16,
        colorG: 185,
        colorB: 129,
        colorA: 0.0,
        scale: 1.1,
        duration: 2,
        ease: 'power2.inOut'
      }, 11)
      .to('#dynamic-cursor', { x: 0, y: -20, duration: 2, ease: 'power2.inOut' }, 11)

      // --- PHASE 3: EXCEL FORMULA DIAGNOSTIC ---
      .addLabel('excel', 13)
      .to('#layer-excel', { opacity: 1, x: 0, y: 0, z: 0, rotationY: 0, duration: 2, ease: 'power3.out' }, 13)
      .to('#layer-excel', { pointerEvents: 'auto', duration: 0.1 }, 13)
      .to('.excel-error-cell', {
        x: 'random(-2, 2)',
        y: 'random(-2, 2)',
        repeat: 5,
        yoyo: true,
        duration: 0.08
      }, 14)
      .call(() => {
        // Trigger simulated repair inside scroll
        const repairBtn = document.getElementById('btn-excel-repair');
        if (repairBtn) {
          gsap.to(repairBtn, { scale: 1.06, duration: 0.3, yoyo: true, repeat: 1 });
        }
      }, null, 14.5)
      .call(() => { toggleExcelCells(true); }, null, 15.2) // Bind Excel cells update directly to scrub timeline event
      .to({}, { duration: 1.5 }) // Hold frame

      // Transition from Excel to PPT visual deck
      .to('#layer-excel', { opacity: 0, y: 150, z: -350, rotationX: -15, duration: 2, ease: 'power2.in' }, 17)
      .to('#layer-excel', { pointerEvents: 'none', duration: 0.1 }, 17)

      // Blend cursor morph properties smoothly to projector lines (purple)
      .to(cursorState, {
        radius: 0,
        width: 0,
        height: 0,
        rectAlpha: 0,
        rayLength: 110,
        rayAlpha: 0.9,
        innerRadius: 0,
        orbitRadius: 0,
        orbitAlpha: 0,
        colorR: 139,
        colorG: 92,
        colorB: 246,
        colorA: 0.0,
        scale: 0.9,
        duration: 2,
        ease: 'power2.inOut'
      }, 17)
      .to('#dynamic-cursor', { x: 0, y: 40, duration: 2, ease: 'power2.inOut' }, 17)

      // --- PHASE 4: PPT GENERATION ---
      .addLabel('ppt', 19)
      .to('#layer-ppt', { opacity: 1, z: 0, duration: 2, ease: 'power3.out' }, 19)
      .to('#layer-ppt', { pointerEvents: 'auto', duration: 0.1 }, 19)
      .to({}, { duration: 2.0 }) // Hold frame

      // Transition from PPT to Unified Workbench
      .to('#layer-ppt', { opacity: 0, scale: 0.75, z: -300, duration: 2, ease: 'power2.inOut' }, 21)
      .to('#layer-ppt', { pointerEvents: 'none', duration: 0.1 }, 21)

      // Blend cursor morph smoothly to portal spinning rings (indigo)
      .to(cursorState, {
        radius: 0,
        width: 0,
        height: 0,
        rectAlpha: 0,
        rayLength: 0,
        rayAlpha: 0,
        innerRadius: 40,
        orbitRadius: 70,
        orbitAlpha: 0.9,
        colorR: 79,
        colorG: 70,
        colorB: 229,
        colorA: 0.0,
        scale: 0.7,
        duration: 2,
        ease: 'power2.inOut'
      }, 21)
      .to('#dynamic-cursor', { x: 0, y: 0, duration: 2, ease: 'power2.inOut' }, 21)

      // --- PHASE 5: UNIFIED CONVERGED WORKBENCH ---
      .addLabel('desktop', 23)
      .to('#layer-desktop', { opacity: 1, z: 0, duration: 2, ease: 'power3.out' }, 23)
      .to('#layer-desktop', { pointerEvents: 'auto', duration: 0.1 }, 23)
      .to({}, { duration: 2.0 }) // Hold frame

      // Transition from Workbench to final Portal zoom-in
      .to('#layer-desktop', { scale: 3, z: 800, opacity: 0, duration: 2.5, ease: 'power2.in' }, 25)
      .to('#layer-desktop', { pointerEvents: 'none', duration: 0.1 }, 25)

      // Morph smoothly back to celestial spark dot
      .to(cursorState, {
        radius: 50,
        width: 0,
        height: 0,
        rectAlpha: 0,
        rayLength: 0,
        rayAlpha: 0,
        innerRadius: 0,
        orbitRadius: 0,
        orbitAlpha: 0,
        colorR: 59,
        colorG: 130,
        colorB: 246,
        colorA: 0.8,
        scale: 2.5,
        rotation: Math.PI * 4,
        duration: 2.5,
        ease: 'power2.inOut'
      }, 25)

      // --- PHASE 6: FINAL DOWNLOAD CTA PORTAL ---
      .addLabel('portal', 27.5)
      .to('#layer-portal', { opacity: 1, z: 0, duration: 2, ease: 'power3.out' }, 27.5)
      .to('#layer-portal', { pointerEvents: 'auto', duration: 0.1 }, 27.5)
      .to(cursorState, {
        scale: 0.0,
        duration: 1.5,
        ease: 'power3.inOut'
      }, 27.5)
      .to('#dynamic-cursor', { x: 0, y: -180, duration: 1.5, ease: 'power3.inOut' }, 27.5);

    // ==================== MANUAL ACCEPT BUTTON ACTION SIMULATION ====================
    document.getElementById('btn-revision-accept').addEventListener('click', () => {
      // Simulate transition into final clean document text
      gsap.to('#layer-word-revision', { opacity: 0, y: 230, duration: 0.4 });

      const delEl = document.getElementById('word-diff-del');
      const addEl = document.getElementById('word-diff-add');

      gsap.to(delEl, { height: 0, py: 0, border: 0, margin: 0, opacity: 0, duration: 0.5 });
      gsap.to(addEl, {
        backgroundColor: 'transparent',
        borderColor: 'transparent',
        color: '#334155', // Normal text slate-700
        fontWeight: 'normal',
        fontSize: '12px',
        padding: 0,
        margin: 0,
        textContent: 'H1财年核心产品线总营收同比增长达 23.5%；通过深度用户回访测试，客户满意度大幅攀升至 96.3%（NPS 增长至 +12），开创历史新高。',
        duration: 0.6,
        delay: 0.2
      });
    });
"""

if old_timeline_block not in content:
    old_timeline_block = content[content.find("// Typewriter effect simulator helper") : content.find("</script>")]

content = content.replace(old_timeline_block, new_timeline_block)

with open('website/index.html', 'w', encoding='utf-8') as f:
    f.write(content)

print("index.html refactored successfully.")

(() => {
  'use strict';

  const slides = Array.from(document.querySelectorAll('.slide'));
  const viewport = document.querySelector('.viewport');
  const picker = document.getElementById('slide-picker');
  const previous = document.getElementById('previous');
  const next = document.getElementById('next');
  const counter = document.getElementById('counter');
  const notesDialog = document.getElementById('notes-dialog');
  const overviewDialog = document.getElementById('overview-dialog');
  const imageDialog = document.getElementById('image-dialog');
  const announcement = document.getElementById('announcement');
  let current = 0;
  let touchStart;

  const announce = (message) => { announcement.textContent = message; };
  const hashIndex = () => {
    const match = /^#slide-(\d+)$/.exec(location.hash);
    return match ? Math.max(0, Math.min(slides.length - 1, Number(match[1]) - 1)) : 0;
  };

  function fit() {
    const scale = Math.min((viewport.clientWidth - 24) / 1440, (viewport.clientHeight - 20) / 810, 1.4);
    document.documentElement.style.setProperty('--scale', String(Math.max(.1, scale)));
  }

  function showSlide(index, updateHash = true) {
    current = Math.max(0, Math.min(slides.length - 1, index));
    slides.forEach((slide, i) => { slide.hidden = i !== current; });
    picker.value = String(current);
    previous.disabled = current === 0;
    next.disabled = current === slides.length - 1;
    counter.textContent = `${String(current + 1).padStart(2, '0')} / ${slides.length}`;
    document.getElementById('progress-fill').style.width = `${(current + 1) / slides.length * 100}%`;
    document.title = `${current + 1}. ${slides[current].dataset.title} | Today intranet`;
    document.querySelectorAll('#overview-list button').forEach((button, i) => {
      button.setAttribute('aria-current', String(i === current));
    });
    if (updateHash && location.hash !== `#slide-${current + 1}`) location.hash = `slide-${current + 1}`;
    announce(`Slide ${current + 1} of ${slides.length}: ${slides[current].dataset.title}`);
    if (matchMedia('(max-width: 760px)').matches) window.scrollTo(0, 0);
  }

  slides.forEach((slide, index) => {
    slide.id = `slide-${index + 1}`;
    slide.dataset.number = String(index + 1).padStart(2, '0');
    slide.setAttribute('aria-label', `${index + 1} of ${slides.length}: ${slide.dataset.title}`);
    const option = document.createElement('option');
    option.value = String(index);
    option.textContent = `${String(index + 1).padStart(2, '0')}  ${slide.dataset.title}`;
    picker.append(option);
    const button = document.createElement('button');
    button.textContent = option.textContent;
    button.addEventListener('click', () => {
      overviewDialog.close();
      showSlide(index);
    });
    document.getElementById('overview-list').append(button);
  });

  function openNotes() {
    document.getElementById('notes-title').textContent = `${current + 1}. ${slides[current].dataset.title}`;
    const content = document.getElementById('notes-content');
    content.replaceChildren();
    const notes = slides[current].querySelector('.notes');
    if (notes) content.append(...Array.from(notes.childNodes).map(node => node.cloneNode(true)));
    notesDialog.showModal();
  }

  async function toggleFullscreen() {
    try {
      if (document.fullscreenElement) await document.exitFullscreen();
      else await document.documentElement.requestFullscreen();
    } catch (error) {
      console.error('Presentation fullscreen request failed.', error);
      announce('Fullscreen is unavailable in this browser. Use the browser fullscreen command instead.');
    }
  }

  document.body.classList.add('enhanced');
  picker.addEventListener('change', () => showSlide(Number(picker.value)));
  previous.addEventListener('click', () => showSlide(current - 1));
  next.addEventListener('click', () => showSlide(current + 1));
  document.getElementById('notes-button').addEventListener('click', openNotes);
  document.getElementById('overview-button').addEventListener('click', () => overviewDialog.showModal());
  document.getElementById('fullscreen-button').addEventListener('click', toggleFullscreen);
  document.getElementById('print-button').addEventListener('click', () => window.print());
  document.querySelectorAll('.dialog-close').forEach(button => {
    button.addEventListener('click', () => button.closest('dialog').close());
  });
  document.querySelectorAll('.shot').forEach(button => {
    button.addEventListener('click', () => {
      const source = button.querySelector('img');
      const target = document.getElementById('expanded-image');
      target.src = source.src;
      target.alt = source.alt;
      imageDialog.showModal();
    });
  });
  document.querySelectorAll('pre').forEach(pre => {
    const button = document.createElement('button');
    button.className = 'copy';
    button.textContent = 'Copy';
    button.setAttribute('aria-label', 'Copy code example');
    button.addEventListener('click', async () => {
      try {
        await navigator.clipboard.writeText(pre.textContent);
        button.textContent = 'Copied';
        announce('Code copied.');
        setTimeout(() => { button.textContent = 'Copy'; }, 1800);
      } catch (error) {
        console.error('Could not copy the presentation example.', error);
        const range = document.createRange();
        range.selectNodeContents(pre);
        const selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(range);
        button.textContent = 'Use Ctrl+C';
        announce('Clipboard access is unavailable. Code is selected; press Ctrl+C or Command+C to copy.');
      }
    });
    pre.parentElement.append(button);
  });
  document.addEventListener('keydown', event => {
    const dialog = document.querySelector('dialog[open]');
    if (dialog) {
      if (event.key === 'Escape') {
        event.preventDefault();
        dialog.close();
      }
      return;
    }
    if (event.altKey || event.ctrlKey || event.metaKey) return;
    if (event.target.closest('input, textarea, select, [contenteditable]')) return;
    if ((event.key === ' ' || event.key === 'Enter') && event.target.closest('button, a')) return;
    switch (event.key.toLowerCase()) {
      case 'arrowright': case 'arrowdown': case 'pagedown': case ' ':
        event.preventDefault(); showSlide(current + 1); break;
      case 'arrowleft': case 'arrowup': case 'pageup':
        event.preventDefault(); showSlide(current - 1); break;
      case 'home': event.preventDefault(); showSlide(0); break;
      case 'end': event.preventDefault(); showSlide(slides.length - 1); break;
      case 'n': openNotes(); break;
      case 'o': overviewDialog.showModal(); break;
      case 'f': toggleFullscreen(); break;
    }
  });
  viewport.addEventListener('touchstart', event => {
    if (event.target.closest('button, a, pre')) return;
    touchStart = { x: event.changedTouches[0].clientX, y: event.changedTouches[0].clientY };
  }, { passive: true });
  viewport.addEventListener('touchend', event => {
    if (!touchStart) return;
    const dx = event.changedTouches[0].clientX - touchStart.x;
    const dy = event.changedTouches[0].clientY - touchStart.y;
    touchStart = undefined;
    if (Math.abs(dx) > 80 && Math.abs(dx) > Math.abs(dy) * 2) showSlide(current + (dx < 0 ? 1 : -1));
  }, { passive: true });
  window.addEventListener('hashchange', () => showSlide(hashIndex(), false));
  window.addEventListener('resize', fit);
  new ResizeObserver(fit).observe(viewport);
  showSlide(hashIndex(), false);
  fit();
})();

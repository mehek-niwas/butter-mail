/**
 * Timeline view - email threads as card stacks.
 */
(function () {
  function buildThreads(emails) {
    console.log('[butter-mail] timeline (buildThreads): starting for', emails.length, 'emails');
    const byMessageId = {};
    emails.forEach((e) => {
      const mid = (e.messageId || '').trim();
      if (mid) byMessageId[mid] = e;
    });

    const visited = new Set();
    const linkedThreads = [];

    function collectFrom(email, thread) {
      if (visited.has(email.id)) return;
      visited.add(email.id);
      thread.push(email);
      const refs = (email.references || '').split(/\s+/).filter(Boolean);
      const inReply = (email.inReplyTo || '').split(/\s+/).filter(Boolean);
      [...refs, ...inReply].forEach((mid) => {
        const parent = byMessageId[mid];
        if (parent) collectFrom(parent, thread);
      });
      emails.forEach((other) => {
        const orefs = (other.references || '').split(/\s+/).filter(Boolean);
        const oinReply = (other.inReplyTo || '').split(/\s+/).filter(Boolean);
        const mid = (email.messageId || '').trim();
        if (mid && [...orefs, ...oinReply].includes(mid)) collectFrom(other, thread);
      });
    }

    emails.forEach((email) => {
      if (visited.has(email.id)) return;
      const thread = [];
      collectFrom(email, thread);
      thread.sort((a, b) => new Date(a.date || 0) - new Date(b.date || 0));
      linkedThreads.push(thread);
    });

    const multiEmail = linkedThreads.filter((t) => t.length > 1);
    const singleEmail = linkedThreads.filter((t) => t.length === 1);

    const bySubject = {};
    singleEmail.forEach((t) => {
      const email = t[0];
      const key = normalizeSubject(email.subject || '(no subject)');
      if (!bySubject[key]) bySubject[key] = [];
      bySubject[key].push(email);
    });
    const subjectThreads = Object.values(bySubject).map((arr) => {
      arr.sort((a, b) => new Date(a.date || 0) - new Date(b.date || 0));
      return arr;
    });

    const threads = [...multiEmail, ...subjectThreads];
    threads.sort((a, b) => {
      const da = new Date((a[a.length - 1] || {}).date || 0);
      const db = new Date((b[b.length - 1] || {}).date || 0);
      return db - da;
    });

    console.log('[butter-mail] timeline (buildThreads): done. threads:', threads.length, 'multi-email:', multiEmail.length, 'by-subject:', subjectThreads.length);
    return threads;
  }

  function normalizeSubject(s) {
    return s.replace(/^(re:\s*|fwd:\s*|fw:\s*)+/gi, '').trim().toLowerCase();
  }

  function render(containerId, emails, onEmailClick, getCategoryColor) {
    const container = document.getElementById(containerId || 'timeline-stacks');
    if (!container) return;
    const getColor = typeof getCategoryColor === 'function' ? getCategoryColor : () => '#B8952E';

    console.log('[butter-mail] timeline (render): building and rendering threads for', emails.length, 'emails');
    const threads = buildThreads(emails);
    container.innerHTML = '';

    threads.forEach((thread) => {
      const stack = document.createElement('div');
      stack.className = 'timeline-stack';
      stack.style.minHeight = (thread.length * 52 + Math.max(0, thread.length - 1) * 8) + 'px';
      thread.forEach((email, ei) => {
        const card = document.createElement('div');
        card.className = 'timeline-card';
        card.style.transform = `translate(${ei * 10}px, ${ei * 10}px)`;
        card.style.zIndex = ei;
        const catColor = (email.categoryId && getColor(email.categoryId)) || getColor(null);
        card.style.borderLeftColor = catColor;
        card.innerHTML = `
          <span class="timeline-card-subject">${escapeHtml(email.subject || '(no subject)')}</span>
          <span class="timeline-card-date">${formatDate(email.date)}</span>
        `;
        card.addEventListener('click', (e) => {
          if (stack.classList.contains('expanded')) {
            e.stopPropagation();
            if (onEmailClick) onEmailClick(email);
          }
        });
        stack.appendChild(card);
      });
      stack.addEventListener('click', () => {
        if (thread.length > 1) {
          stack.classList.toggle('expanded');
        } else if (thread.length === 1 && onEmailClick) {
          onEmailClick(thread[0]);
        }
      });
      container.appendChild(stack);
    });
  }

  function escapeHtml(str) {
    if (!str) return '';
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  function formatDate(dateStr) {
    if (!dateStr) return '';
    try {
      const d = new Date(dateStr);
      if (isNaN(d.getTime())) return dateStr;
      return d.toLocaleDateString(undefined, {
        month: 'short',
        day: 'numeric',
        year: d.getFullYear() !== new Date().getFullYear() ? 'numeric' : undefined
      });
    } catch {
      return dateStr;
    }
  }

  window.TimelineView = {
    buildThreads,
    render
  };
})();

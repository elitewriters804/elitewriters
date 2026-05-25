// ============================================
// ELITEWRITERS — MAIN JAVASCRIPT
// ============================================

// ── Navbar scroll effect ──
const navbar = document.getElementById('navbar');
if (navbar) {
  window.addEventListener('scroll', () => {
    navbar.classList.toggle('scrolled', window.scrollY > 50);
  });
}

// ── Mobile nav toggle ──
const navToggle = document.getElementById('nav-toggle');
const navLinks = document.getElementById('nav-links');
if (navToggle && navLinks) {
  navToggle.addEventListener('click', () => {
    navLinks.classList.toggle('open');
  });
  document.addEventListener('click', (e) => {
    if (!navbar.contains(e.target)) navLinks.classList.remove('open');
  });
}

// ── Price Calculator ──
function initCalculator() {
  const form = document.getElementById('price-calculator');
  if (!form) return;

  const rates = {
    academic: {
      perPage: { highschool: 12, undergraduate: 15, masters: 18, phd: 22 },
      deadlineMult: { '14': 1, '7': 1.2, '3': 1.33, '2': 1.5 },
    },
    blog: { perWord: 0.08 },
    business: { perWord: 0.10 }
  };

  function fmt(n) { return '$' + n.toFixed(2); }

  function calculate() {
    const service = document.getElementById('calc-service').value;
    const level = document.getElementById('calc-level').value;
    const qty = Math.max(1, parseInt(document.getElementById('calc-qty').value) || 1);
    const deadline = document.getElementById('calc-deadline').value;
    const hasPlag = document.getElementById('calc-plagiarism') && document.getElementById('calc-plagiarism').checked;
    const hasProof = document.getElementById('calc-proofreading') && document.getElementById('calc-proofreading').checked;
    const hasFormat = document.getElementById('calc-formatting') && document.getElementById('calc-formatting').checked;

    const levelGroup = document.getElementById('calc-level-group');
    const qtyLabel = document.getElementById('calc-qty-label');
    const baseLabel = document.getElementById('calc-base-label');

    let base = 0;
    let rushAmt = 0;
    let proofAmt = 0;
    const isRush = deadline === '2';

    if (service === 'academic') {
      if (levelGroup) levelGroup.style.display = 'block';
      if (qtyLabel) qtyLabel.textContent = 'Number of pages';
      const rate = rates.academic.perPage[level] || 15;
      const mult = isRush ? 1 : (rates.academic.deadlineMult[deadline] || 1);
      base = rate * qty * mult;
      if (isRush) rushAmt = base * 0.5;
      if (baseLabel) baseLabel.textContent = qty + (qty > 1 ? ' pages' : ' page') + ' × $' + rate + ((!isRush && deadline !== '14') ? ' × ' + mult + 'x deadline' : '');
      proofAmt = hasProof ? 8 * qty : 0;
    } else {
      if (levelGroup) levelGroup.style.display = 'none';
      if (qtyLabel) qtyLabel.textContent = 'Number of words';
      const rate = service === 'blog' ? 0.08 : 0.10;
      base = rate * qty;
      if (isRush) rushAmt = base * 0.5;
      if (baseLabel) baseLabel.textContent = qty + ' words × $' + rate;
      const pages = Math.ceil(qty / 275);
      proofAmt = hasProof ? pages * 8 : 0;
    }

    const total = base + rushAmt + (hasPlag ? 5 : 0) + proofAmt + (hasFormat ? 10 : 0);

    const setEl = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
    const showEl = (id, show) => { const el = document.getElementById(id); if (el) el.style.display = show ? 'flex' : 'none'; };

    setEl('calc-base-price', fmt(base));
    setEl('calc-rush-price', fmt(rushAmt));
    setEl('calc-proof-price', fmt(proofAmt));
    setEl('calc-total', fmt(total));
    showEl('calc-rush-row', isRush);
    showEl('calc-plag-row', hasPlag);
    showEl('calc-proof-row', hasProof);
    showEl('calc-format-row', hasFormat);

    const noteEl = document.getElementById('calc-payment-note');
    if (noteEl) {
      noteEl.textContent = total >= 30
        ? 'Pay $' + (total / 2).toFixed(2) + ' upfront to confirm, then $' + (total / 2).toFixed(2) + ' on delivery.'
        : 'Full payment of ' + fmt(total) + ' required upfront for this order.';
    }
  }

  ['calc-service','calc-level','calc-qty','calc-deadline'].forEach(id => {
    const el = document.getElementById(id);
    if (el) { el.addEventListener('input', calculate); el.addEventListener('change', calculate); }
  });
  ['calc-plagiarism','calc-proofreading','calc-formatting'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('change', calculate);
  });

  calculate();
}

// ── Scroll animations ──
function initScrollAnimations() {
  const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.style.opacity = '1';
        entry.target.style.transform = 'translateY(0)';
      }
    });
  }, { threshold: 0.1 });

  document.querySelectorAll('.card, .testimonial-card').forEach(el => {
    el.style.opacity = '0';
    el.style.transform = 'translateY(20px)';
    el.style.transition = 'opacity 0.6s ease, transform 0.6s ease';
    observer.observe(el);
  });
}

// ── Init on DOM ready ──
document.addEventListener('DOMContentLoaded', () => {
  initCalculator();
  initScrollAnimations();
});

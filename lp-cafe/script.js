// ===== ヘッダーの背景・トップへ戻るボタン =====
const header = document.querySelector('.header');
const toTop = document.querySelector('.to-top');
const onScroll = () => {
  const y = window.scrollY;
  header.classList.toggle('is-scrolled', y > 40);
  toTop.classList.toggle('is-show', y > 600);
};
window.addEventListener('scroll', onScroll, { passive: true });
onScroll();

// ===== スマホ用メニュー =====
const menuBtn = document.querySelector('.menu-btn');
const nav = document.getElementById('nav');
const setMenu = (open) => {
  menuBtn.setAttribute('aria-expanded', String(open));
  menuBtn.setAttribute('aria-label', open ? 'メニューを閉じる' : 'メニューを開く');
  nav.classList.toggle('is-open', open);
  document.body.style.overflow = open ? 'hidden' : '';
};
menuBtn.addEventListener('click', () => setMenu(menuBtn.getAttribute('aria-expanded') !== 'true'));
nav.querySelectorAll('a').forEach((a) => a.addEventListener('click', () => setMenu(false)));
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') setMenu(false); });

// ===== スクロールで要素をふわっと表示 =====
const io = new IntersectionObserver((entries) => {
  entries.forEach((entry) => {
    if (entry.isIntersecting) {
      entry.target.classList.add('is-visible');
      io.unobserve(entry.target);
    }
  });
}, { threshold: 0.15 });
document.querySelectorAll('.fade').forEach((el) => io.observe(el));

// ===== メニューのタブ切り替え =====
const tabs = [...document.querySelectorAll('.tab')];
const selectTab = (tab) => {
  tabs.forEach((t) => {
    const active = t === tab;
    t.classList.toggle('is-active', active);
    t.setAttribute('aria-selected', String(active));
    t.tabIndex = active ? 0 : -1;
    const panel = document.getElementById(t.getAttribute('aria-controls'));
    panel.hidden = !active;
    panel.classList.toggle('is-active', active);
  });
};
tabs.forEach((tab, i) => {
  tab.addEventListener('click', () => selectTab(tab));
  tab.addEventListener('keydown', (e) => {
    const dir = e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0;
    if (!dir) return;
    const next = tabs[(i + dir + tabs.length) % tabs.length];
    selectTab(next);
    next.focus();
  });
});

// ===== 予約フォームの入力チェック =====
const form = document.getElementById('reserveForm');
const dateInput = document.getElementById('date');

// 明日以降・水曜定休を考慮
const toISO = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const tomorrow = new Date();
tomorrow.setDate(tomorrow.getDate() + 1);
dateInput.min = toISO(tomorrow);

const messages = {
  name: (v) => (!v.trim() ? 'お名前を入力してください' : ''),
  email: (v) => (!v.trim() ? 'メールアドレスを入力してください'
    : !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v) ? 'メールアドレスの形式が正しくありません' : ''),
  date: (v) => {
    if (!v) return 'ご来店日を選択してください';
    if (v < dateInput.min) return 'ご予約は明日以降の日付でお願いします';
    const d = new Date(`${v}T00:00:00`);
    if (d.getDay() === 3) return '水曜日は定休日です';
    if (d.getDay() === 2 && Math.ceil(d.getDate() / 7) === 3) return '第3火曜日は定休日です';
    return '';
  },
  time: (v) => (!v ? '時間を選択してください' : ''),
  people: (v) => (!v ? '人数を選択してください' : ''),
};

const validate = (field) => {
  const check = messages[field.name];
  if (!check) return true;
  const msg = check(field.value);
  field.classList.toggle('is-invalid', !!msg);
  field.setAttribute('aria-invalid', String(!!msg));
  field.closest('.form__row').querySelector('.form__error').textContent = msg;
  return !msg;
};

Object.keys(messages).forEach((name) => {
  const field = form.elements[name];
  field.addEventListener('blur', () => validate(field));
  field.addEventListener('change', () => validate(field));
});

form.addEventListener('submit', (e) => {
  e.preventDefault();
  const fields = Object.keys(messages).map((n) => form.elements[n]);
  const results = fields.map(validate);
  const firstInvalid = fields[results.indexOf(false)];
  if (firstInvalid) {
    firstInvalid.focus();
    return;
  }
  document.getElementById('formDone').hidden = false;
  form.reset();
});
form.addEventListener('input', () => { document.getElementById('formDone').hidden = true; });

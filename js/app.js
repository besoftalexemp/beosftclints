(() => {
  'use strict';

  const EXCEL_PATH = 'data/customers.xlsx';
  const PAGE_SIZE = 24;
  const REQUIRED_HEADERS = ['المحافظة', 'اسم العميل', 'العنوان'];

  const state = {
    customers: [],
    filtered: [],
    counts: new Map(),
    selectedGovernorate: '',
    searchTerm: '',
    visibleCount: PAGE_SIZE,
  };

  const governorates = [
    { name: 'الاسكندرية', x: 36, y: 12 },
    { name: 'البحيرة', x: 42, y: 16 },
    { name: 'كفر الشيخ', x: 48, y: 13 },
    { name: 'الدقهلية', x: 56, y: 15 },
    { name: 'دمياط', x: 63, y: 13 },
    { name: 'الغربية', x: 50, y: 19 },
    { name: 'المنوفية', x: 46, y: 22 },
    { name: 'القليوبية', x: 54, y: 23 },
    { name: 'الشرقية', x: 62, y: 21 },
    { name: 'بورسعيد', x: 72, y: 17 },
    { name: 'الإسماعيلية', x: 70, y: 24 },
    { name: 'السويس', x: 70, y: 30 },
    { name: 'القاهرة', x: 57, y: 29 },
    { name: 'الجيزة', x: 48, y: 31 },
    { name: 'الفيوم', x: 42, y: 38 },
    { name: 'بني سويف', x: 48, y: 44 },
    { name: 'المنيا', x: 48, y: 52 },
    { name: 'أسيوط', x: 49, y: 61 },
    { name: 'سوهاج', x: 51, y: 69 },
    { name: 'قنا', x: 58, y: 76 },
    { name: 'الأقصر', x: 60, y: 82 },
    { name: 'أسوان', x: 63, y: 91 },
    { name: 'البحر الأحمر', x: 74, y: 63 },
    { name: 'الوادي الجديد', x: 29, y: 66 },
    { name: 'مطروح', x: 18, y: 20 },
    { name: 'شمال سيناء', x: 84, y: 20 },
    { name: 'جنوب سيناء', x: 82, y: 36 },
  ];

  const aliases = {
    'اسكندريه': 'الاسكندرية',
    'الاسكندريه': 'الاسكندرية',
    'الاسكندرية': 'الاسكندرية',

    'قاهره': 'القاهرة',
    'القاهره': 'القاهرة',
    'القاهرة': 'القاهرة',

    'جيزه': 'الجيزة',
    'الجيزه': 'الجيزة',
    'الجيزة': 'الجيزة',

    'شرقيه': 'الشرقية',
    'الشرقيه': 'الشرقية',
    'الشرقية': 'الشرقية',

    'غربيه': 'الغربية',
    'الغربيه': 'الغربية',
    'الغربية': 'الغربية',

    'دقهليه': 'الدقهلية',
    'الدقهليه': 'الدقهلية',
    'الدقهلية': 'الدقهلية',

    'منوفيه': 'المنوفية',
    'المنوفيه': 'المنوفية',
    'المنوفية': 'المنوفية',

    'قليوبيه': 'القليوبية',
    'القليوبيه': 'القليوبية',
    'القليوبية': 'القليوبية',

    'بحيره': 'البحيرة',
    'البحيره': 'البحيرة',
    'البحيرة': 'البحيرة',

    'منيا': 'المنيا',
    'المنيا': 'المنيا',

    'سوهاج': 'سوهاج',

    'اسيوط': 'أسيوط',
    'أسيوط': 'أسيوط',

    'اسوان': 'أسوان',
    'أسوان': 'أسوان',

    'قنا': 'قنا',

    'اقصر': 'الأقصر',
    'الاقصر': 'الأقصر',
    'الأقصر': 'الأقصر',

    'فيوم': 'الفيوم',
    'الفيوم': 'الفيوم',

    'بني سويف': 'بني سويف',
    'بنى سويف': 'بني سويف',

    'كفر الشيخ': 'كفر الشيخ',
    'دمياط': 'دمياط',

    'بورسعيد': 'بورسعيد',
    'بور سعيد': 'بورسعيد',

    'اسماعيليه': 'الإسماعيلية',
    'الاسماعيليه': 'الإسماعيلية',
    'الإسماعيلية': 'الإسماعيلية',

    'سويس': 'السويس',
    'السويس': 'السويس',

    'مطروح': 'مطروح',

    'وادي جديد': 'الوادي الجديد',
    'الوادي الجديد': 'الوادي الجديد',

    'بحر احمر': 'البحر الأحمر',
    'البحر الاحمر': 'البحر الأحمر',
    'البحر الأحمر': 'البحر الأحمر',

    'شمال سيناء': 'شمال سيناء',
    'جنوب سيناء': 'جنوب سيناء',
  };

  const els = {
    totalCustomers: document.getElementById('totalCustomers'),
    totalGovernorates: document.getElementById('totalGovernorates'),
    visibleAddresses: document.getElementById('visibleAddresses'),

    selectedTitle: document.getElementById('selectedGovernorateTitle'),
    selectedCount: document.getElementById('selectedGovernorateCount'),
    selectedText: document.getElementById('selectedGovernorateText'),

    topList: document.getElementById('topGovernoratesList'),
    mapMarkers: document.getElementById('mapMarkers'),
    mapFocus: document.getElementById('mapFocus'),
    mobileMapList: document.getElementById('mobileMapList'),

    showAllBtn: document.getElementById('showAllBtn'),

    governorateSelect: document.getElementById('governorateSelect'),
    searchInput: document.getElementById('searchInput'),
    resetFilters: document.getElementById('resetFilters'),

    resultsCount: document.getElementById('resultsCount'),
    grid: document.getElementById('customersGrid'),

    loading: document.getElementById('loadingState'),
    error: document.getElementById('errorState'),

    excelFileInput: document.getElementById('excelFileInput'),
    loadMore: document.getElementById('loadMoreBtn'),

  };

  function cleanText(value) {
    if (value === null || value === undefined) return '';

    return String(value)
      .replace(/\s+/g, ' ')
      .trim();
  }

  function normalizeArabic(value) {
    return cleanText(value)
      .replace(/[ًٌٍَُِّْـ]/g, '')
      .replace(/[أإآ]/g, 'ا')
      .replace(/ى/g, 'ي')
      .replace(/ة/g, 'ه')
      .replace(/[\-–—_/\\]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function canonicalGovernorate(value) {
    const raw = cleanText(value);

    if (!raw) {
      return '';
    }

    const directKey = normalizeArabic(raw);
    const keyWithoutAl = directKey.replace(/^ال/, '');

    for (const [alias, canonical] of Object.entries(aliases)) {
      const aliasNorm = normalizeArabic(alias);
      const aliasWithoutAl = aliasNorm.replace(/^ال/, '');

      if (
        directKey === aliasNorm ||
        keyWithoutAl === aliasWithoutAl ||
        directKey.includes(aliasNorm) ||
        aliasNorm.includes(directKey)
      ) {
        return canonical;
      }
    }

    const match = governorates.find((gov) => {
      const govNorm = normalizeArabic(gov.name);
      const govWithoutAl = govNorm.replace(/^ال/, '');

      return (
        directKey === govNorm ||
        keyWithoutAl === govWithoutAl ||
        directKey.includes(govNorm)
      );
    });

    return match ? match.name : raw;
  }

  function locateDataSheet(workbook) {
    const preferredSheet = workbook.SheetNames.includes('2')
      ? '2'
      : null;

    const candidates = preferredSheet
      ? [
          preferredSheet,
          ...workbook.SheetNames.filter(
            (name) => name !== preferredSheet
          ),
        ]
      : workbook.SheetNames;

    for (const sheetName of candidates) {
      const worksheet = workbook.Sheets[sheetName];

      const rows = XLSX.utils.sheet_to_json(
        worksheet,
        {
          header: 1,
          defval: '',
        }
      );

      if (!rows.length) {
        continue;
      }

      const headers = rows[0].map(cleanText);

      const hasRequiredHeaders =
        REQUIRED_HEADERS.every((header) =>
          headers.includes(header)
        );

      if (hasRequiredHeaders) {
        return {
          sheetName,
          worksheet,
        };
      }
    }

    throw new Error(
      'لم يتم العثور على شيت يحتوي الأعمدة المطلوبة: المحافظة، اسم العميل، العنوان'
    );
  }

  function parseWorkbook(workbook) {
    const { worksheet } =
      locateDataSheet(workbook);

    const rows =
      XLSX.utils.sheet_to_json(
        worksheet,
        {
          defval: '',
        }
      );

    return rows
      .map((row, index) => {
        const name =
          cleanText(row['اسم العميل']);

        const governorateRaw =
          cleanText(row['المحافظة']);

        const address =
          cleanText(row['العنوان']);

        if (!name || !governorateRaw) {
          return null;
        }

        return {
          id:
            cleanText(row['م']) ||
            index + 1,

          name,

          governorate:
            canonicalGovernorate(
              governorateRaw
            ),

          governorateRaw,

          address,
        };
      })
      .filter(Boolean);
  }

  async function loadExcelFromServer() {
    if (typeof XLSX === 'undefined') {
      throw new Error(
        'مكتبة قراءة Excel لم يتم تحميلها.'
      );
    }

    const response = await fetch(
      `${EXCEL_PATH}?v=${Date.now()}`
    );

    if (!response.ok) {
      throw new Error(
        `HTTP ${response.status}`
      );
    }

    const data =
      await response.arrayBuffer();

    const workbook =
      XLSX.read(data, {
        type: 'array',
      });

    return parseWorkbook(workbook);
  }

  function readUserExcel(file) {
    return new Promise(
      (resolve, reject) => {
        if (typeof XLSX === 'undefined') {
          reject(
            new Error(
              'مكتبة Excel غير متاحة.'
            )
          );

          return;
        }

        const reader =
          new FileReader();

        reader.onload = (event) => {
          try {
            const workbook =
              XLSX.read(
                event.target.result,
                {
                  type: 'array',
                }
              );

            resolve(
              parseWorkbook(workbook)
            );
          } catch (error) {
            reject(error);
          }
        };

        reader.onerror = reject;

        reader.readAsArrayBuffer(file);
      }
    );
  }

  function buildCounts() {
    state.counts = new Map();

    state.customers.forEach(
      (customer) => {
        const key =
          customer.governorate ||
          'غير محدد';

        state.counts.set(
          key,
          (state.counts.get(key) || 0) + 1
        );
      }
    );
  }

  function formatNumber(number) {
    return new Intl.NumberFormat(
      'ar-EG'
    ).format(number || 0);
  }

  function getTopEntries(limit = 7) {
    return [...state.counts.entries()]
      .sort(
        (a, b) => b[1] - a[1]
      )
      .slice(0, limit);
  }

  function updateMapFocus(name) {
    if (!els.mapFocus) {
      return;
    }

    if (name) {
      const count =
        state.counts.get(name) || 0;

      els.mapFocus.textContent =
        `${name}: ${formatNumber(count)} عميل`;

      return;
    }

    const top =
      getTopEntries(1)[0];

    els.mapFocus.textContent = top
      ? `أعلى محافظة حاليا: ${top[0]} (${formatNumber(top[1])} عميل)`
      : 'اختر محافظة من الخريطة أو القائمة';
  }

  function renderStats() {
    if (els.totalCustomers) {
      els.totalCustomers.textContent =
        formatNumber(
          state.customers.length
        );
    }

    if (els.totalGovernorates) {
      els.totalGovernorates.textContent =
        formatNumber(
          state.counts.size
        );
    }

    if (els.visibleAddresses) {
      els.visibleAddresses.textContent =
        formatNumber(
          state.customers.filter(
            (customer) =>
              customer.address
          ).length
        );
    }
  }

  function renderTopGovernorates() {
    if (!els.topList) {
      return;
    }

    const top =
      getTopEntries(7);

    const max =
      top[0]?.[1] || 1;

    els.topList.innerHTML =
      top
        .map(
          ([name, count]) => `
            <button
              class="top-row"
              type="button"
              data-gov="${escapeAttr(name)}"
              aria-label="عرض عملاء ${escapeAttr(name)}"
            >
              <span>${escapeHtml(name)}</span>

              <span class="bar">
                <i
                  style="
                    width:${Math.max(
                      8,
                      (count / max) * 100
                    )}%
                  "
                ></i>
              </span>

              <strong>
                ${formatNumber(count)}
              </strong>
            </button>
          `
        )
        .join('');

    els.topList
      .querySelectorAll('[data-gov]')
      .forEach((button) => {
        button.addEventListener(
          'click',
          () => {
            selectGovernorate(
              button.dataset.gov
            );
          }
        );
      });
  }

  function renderGovernorateSelect() {
    if (!els.governorateSelect) {
      return;
    }

    const names =
      [...state.counts.keys()]
        .sort((a, b) =>
          a.localeCompare(
            b,
            'ar'
          )
        );

    els.governorateSelect.innerHTML =
      `
        <option value="">
          كل المحافظات
        </option>
      ` +
      names
        .map(
          (name) => `
            <option
              value="${escapeAttr(name)}"
            >
              ${escapeHtml(name)}
              (${formatNumber(
                state.counts.get(name)
              )})
            </option>
          `
        )
        .join('');
  }

  function renderMap() {
    if (!els.mapMarkers) {
      return;
    }

    const maxCount =
      Math.max(
        1,
        ...state.counts.values()
      );

    els.mapMarkers.innerHTML =
      governorates
        .map((gov) => {
          const count =
            state.counts.get(
              gov.name
            ) || 0;
          const strength =
            count / maxCount;
          const scale =
            count
              ? 1 + (strength * 0.38)
              : 0.82;

          return `
            <button
              type="button"
              class="
                map-marker
                ${count ? 'has-data' : 'zero'}
              "
              data-gov="${escapeAttr(
                gov.name
              )}"
              data-count="${count}"
              style="
                left:${gov.x}%;
                top:${gov.y}%;
                --pin-scale:${scale.toFixed(2)};
                --pin-strength:${(0.12 + strength * 0.32).toFixed(2)};
              "
              title="${escapeAttr(
                gov.name
              )}: ${count} عميل"
              aria-label="${escapeAttr(
                gov.name
              )}: ${formatNumber(count)} عميل"
            >
              <span class="pin-dot" aria-hidden="true"></span>
              <span class="pin-label">
                ${escapeHtml(gov.name)}
                <small>${formatNumber(count)} عميل</small>
              </span>
            </button>
          `;
        })
        .join('');

    els.mapMarkers
      .querySelectorAll(
        '.map-marker'
      )
      .forEach((marker) => {
        marker.addEventListener(
          'click',
          () => {
            selectGovernorate(
              marker.dataset.gov
            );
          }
        );

        marker.addEventListener(
          'pointerenter',
          () => updateMapFocus(
            marker.dataset.gov
          )
        );

        marker.addEventListener(
          'focus',
          () => updateMapFocus(
            marker.dataset.gov
          )
        );

        marker.addEventListener(
          'pointerleave',
          () => updateMapFocus(
            state.selectedGovernorate
          )
        );

        marker.addEventListener(
          'blur',
          () => updateMapFocus(
            state.selectedGovernorate
          )
        );
      });

    renderMobileMapList();

    updateMapFocus(
      state.selectedGovernorate
    );
  }

  function renderMobileMapList() {
    if (!els.mobileMapList) {
      return;
    }

    const ordered =
      [...governorates]
        .sort((a, b) => {
          const countDiff =
            (state.counts.get(b.name) || 0) -
            (state.counts.get(a.name) || 0);

          if (countDiff) {
            return countDiff;
          }

          return a.name.localeCompare(
            b.name,
            'ar'
          );
        });

    els.mobileMapList.innerHTML =
      ordered
        .map((gov) => {
          const count =
            state.counts.get(gov.name) || 0;

          return `
            <button
              type="button"
              data-gov="${escapeAttr(gov.name)}"
              aria-label="عرض عملاء ${escapeAttr(gov.name)}"
            >
              ${escapeHtml(gov.name)}
              <span>${formatNumber(count)}</span>
            </button>
          `;
        })
        .join('');

    els.mobileMapList
      .querySelectorAll('[data-gov]')
      .forEach((button) => {
        button.addEventListener(
          'click',
          () => selectGovernorate(
            button.dataset.gov
          )
        );
      });
  }

  function updateMapActive() {
    if (els.mapMarkers) {
      els.mapMarkers
        .querySelectorAll(
          '.map-marker'
        )
        .forEach((marker) => {
          marker.classList.toggle(
            'active',
            marker.dataset.gov ===
              state.selectedGovernorate
          );
        });
    }

    document
      .querySelectorAll('[data-gov]')
      .forEach((item) => {
        item.classList.toggle(
          'active',
          item.dataset.gov ===
            state.selectedGovernorate
        );
      });

    updateMapFocus(
      state.selectedGovernorate
    );
  }

  function updateCoveragePanel() {
    if (
      !els.selectedTitle ||
      !els.selectedCount ||
      !els.selectedText
    ) {
      return;
    }

    if (state.selectedGovernorate) {
      const count =
        state.counts.get(
          state.selectedGovernorate
        ) || 0;

      els.selectedTitle.textContent =
        state.selectedGovernorate;

      els.selectedCount.textContent =
        formatNumber(count);

      els.selectedText.textContent =
        count
          ? `يعرض الدليل بالأسفل العملاء المسجلين في ${state.selectedGovernorate}.`
          : 'لا توجد سجلات مطابقة لهذه المحافظة في ملف Excel الحالي.';
    } else {
      els.selectedTitle.textContent =
        'كل المحافظات';

      els.selectedCount.textContent =
        formatNumber(
          state.customers.length
        );

      els.selectedText.textContent =
        'اختر محافظة من الخريطة أو من قائمة الفلترة لعرض العملاء المرتبطين بها.';
    }
  }

  function selectGovernorate(name) {
    state.selectedGovernorate =
      name || '';

    state.visibleCount =
      PAGE_SIZE;

    if (els.governorateSelect) {
      els.governorateSelect.value =
        state.selectedGovernorate;
    }

    updateMapActive();

    updateCoveragePanel();

    applyFilters();
  }

  function applyFilters() {
    const needle =
      normalizeArabic(
        state.searchTerm
      );

    state.filtered =
      state.customers.filter(
        (customer) => {
          const governorateMatches =
            !state.selectedGovernorate ||
            customer.governorate ===
              state.selectedGovernorate;

          if (!governorateMatches) {
            return false;
          }

          if (!needle) {
            return true;
          }

          const haystack =
            normalizeArabic(
              `${customer.name} ${customer.address} ${customer.governorate}`
            );

          return haystack.includes(
            needle
          );
        }
      );

    renderCustomers();
  }

  function renderCustomers() {
    if (!els.grid) {
      return;
    }

    const shown =
      state.filtered.slice(
        0,
        state.visibleCount
      );

    if (els.resultsCount) {
      els.resultsCount.textContent =
        `النتائج: ${formatNumber(
          state.filtered.length
        )} عميل`;
    }

    if (!shown.length) {
      els.grid.innerHTML = `
        <div class="empty-state">
          <strong>
            لا توجد نتائج مطابقة
          </strong>

          <br>

          جرّب تغيير المحافظة
          أو البحث بكلمة أخرى.
        </div>
      `;

      if (els.loadMore) {
        els.loadMore.classList.add(
          'hidden'
        );
      }

      return;
    }

    els.grid.innerHTML =
      shown
        .map(
          (customer, index) => `
            <article
              class="customer-card"
            >
              <div
                class="customer-index"
              >
                ${formatNumber(
                  index + 1
                )}
              </div>

              <h3>
                ${escapeHtml(
                  customer.name
                )}
              </h3>

              <div
                class="customer-meta"
              >
                <span>
                  📍
                  ${escapeHtml(
                    customer.governorate
                  )}
                </span>
              </div>

              <p
                class="customer-address"
              >
                ${
                  customer.address
                    ? `العنوان: ${escapeHtml(
                        customer.address
                      )}`
                    : 'العنوان غير مسجل في الملف'
                }
              </p>
            </article>
          `
        )
        .join('');

    if (els.loadMore) {
      els.loadMore.classList.toggle(
        'hidden',
        state.visibleCount >=
          state.filtered.length
      );
    }
  }

  function escapeHtml(value) {
    return String(value).replace(
      /[&<>'"]/g,
      (char) =>
        ({
          '&': '&amp;',
          '<': '&lt;',
          '>': '&gt;',
          "'": '&#39;',
          '"': '&quot;',
        })[char]
    );
  }

  function escapeAttr(value) {
    return escapeHtml(value);
  }

  function finishLoading(customers) {
    state.customers = customers;

    state.filtered = customers;

    buildCounts();

    renderStats();

    renderTopGovernorates();

    renderGovernorateSelect();

    renderMap();

    updateCoveragePanel();

    updateMapActive();

    applyFilters();

    if (els.loading) {
      els.loading.classList.add(
        'hidden'
      );
    }

    if (els.error) {
      els.error.classList.add(
        'hidden'
      );
    }
  }

  function showError(error) {
    console.error(error);

    if (els.loading) {
      els.loading.classList.add(
        'hidden'
      );
    }

    if (els.error) {
      els.error.classList.remove(
        'hidden'
      );
    }

    if (els.resultsCount) {
      els.resultsCount.textContent =
        'لم يتم تحميل البيانات';
    }
  }

  if (els.searchInput) {
    els.searchInput.addEventListener(
      'input',
      (event) => {
        state.searchTerm =
          event.target.value;

        state.visibleCount =
          PAGE_SIZE;

        applyFilters();
      }
    );
  }

  if (els.governorateSelect) {
    els.governorateSelect.addEventListener(
      'change',
      (event) => {
        state.selectedGovernorate =
          event.target.value;

        state.visibleCount =
          PAGE_SIZE;

        updateMapActive();

        updateCoveragePanel();

        applyFilters();
      }
    );
  }

  if (els.showAllBtn) {
    els.showAllBtn.addEventListener(
      'click',
      () => {
        selectGovernorate('');
      }
    );
  }

  if (els.resetFilters) {
    els.resetFilters.addEventListener(
      'click',
      () => {
        state.searchTerm = '';

        state.selectedGovernorate =
          '';

        state.visibleCount =
          PAGE_SIZE;

        if (els.searchInput) {
          els.searchInput.value = '';
        }

        if (
          els.governorateSelect
        ) {
          els.governorateSelect.value =
            '';
        }

        updateMapActive();

        updateCoveragePanel();

        applyFilters();
      }
    );
  }

  if (els.loadMore) {
    els.loadMore.addEventListener(
      'click',
      () => {
        state.visibleCount +=
          PAGE_SIZE;

        renderCustomers();
      }
    );
  }

  if (els.excelFileInput) {
    els.excelFileInput.addEventListener(
      'change',
      async (event) => {
        const file =
          event.target.files?.[0];

        if (!file) {
          return;
        }

        if (els.error) {
          els.error.classList.add(
            'hidden'
          );
        }

        if (els.loading) {
          els.loading.classList.remove(
            'hidden'
          );
        }

        try {
          const customers =
            await readUserExcel(
              file
            );

          finishLoading(
            customers
          );
        } catch (error) {
          showError(error);
        }
      }
    );
  }

  function initReveal() {
    const items =
      document.querySelectorAll('[data-reveal]');

    if (!items.length) {
      return;
    }

    if (!('IntersectionObserver' in window)) {
      items.forEach((item) =>
        item.classList.add('is-visible')
      );

      return;
    }

    const observer =
      new IntersectionObserver(
        (entries) => {
          entries.forEach((entry) => {
            if (!entry.isIntersecting) {
              return;
            }

            entry.target.classList.add(
              'is-visible'
            );

            observer.unobserve(
              entry.target
            );
          });
        },
        {
          threshold: 0.12,
          rootMargin: '0px 0px -40px',
        }
      );

    items.forEach((item) =>
      observer.observe(item)
    );
  }

  initReveal();

  loadExcelFromServer()
    .then(finishLoading)
    .catch(showError);
})();

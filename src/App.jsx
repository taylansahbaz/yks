import { useState, useMemo, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import Select from 'react-select';
import { 
  Scale, 
  Cpu, 
  Users, 
  MessageCircle, 
  Building, 
  Briefcase, 
  GraduationCap, 
  Search,
  HardHat,
  Monitor,
  Stethoscope,
  BookOpen,
  Leaf,
  Cog,
  Filter
} from 'lucide-react';
import AdBanner from './components/AdBanner';
import './index.css';

const trToLower = (str) => String(str || '').toLocaleLowerCase('tr-TR');

const getProgramIcon = (fakulteAdi, programAdi) => {
  const fName = String(fakulteAdi || '').toLocaleUpperCase('tr-TR');
  const pName = String(programAdi || '').toLocaleUpperCase('tr-TR');
  
  if (pName.includes('İNŞAAT')) return <HardHat size={48} strokeWidth={1.5} color="var(--color-icon-insaat)" />;
  if (pName.includes('BİLGİSAYAR') || pName.includes('YAZILIM') || pName.includes('BİLİŞİM')) return <Monitor size={48} strokeWidth={1.5} color="var(--color-icon-bilgisayar)" />;
  if (pName.includes('TIP') || pName.includes('DİŞ') || pName.includes('HEMŞİRE') || pName.includes('SAĞLIK')) return <Stethoscope size={48} strokeWidth={1.5} color="var(--color-icon-saglik)" />;
  if (pName.includes('HUKUK')) return <Scale size={48} strokeWidth={1.5} color="var(--color-icon-hukuk)" />;
  if (pName.includes('MİMARLIK')) return <Building size={48} strokeWidth={1.5} color="var(--color-icon-mimarlik)" />;
  if (pName.includes('EĞİTİM') || pName.includes('ÖĞRETMEN') || pName.includes('EDEBİYAT')) return <BookOpen size={48} strokeWidth={1.5} color="var(--color-icon-egitim)" />;
  if (pName.includes('ZİRAAT') || pName.includes('ORMAN') || pName.includes('ÇEVRE')) return <Leaf size={48} strokeWidth={1.5} color="var(--color-icon-ziraat)" />;
  if (pName.includes('MAKİNE') || pName.includes('MEKATRONİK') || pName.includes('ENDÜSTRİ')) return <Cog size={48} strokeWidth={1.5} color="var(--color-icon-makine)" />;
  
  if (fName.includes('MÜHENDİSLİK')) return <Cog size={48} strokeWidth={1.5} color="var(--color-icon-muh)" />;
  if (fName.includes('İNSAN') || fName.includes('TOPLUM') || fName.includes('FEN-EDEBİYAT')) return <Users size={48} strokeWidth={1.5} color="var(--color-icon-insan)" />;
  if (fName.includes('İLETİŞİM')) return <MessageCircle size={48} strokeWidth={1.5} color="var(--color-icon-iletisim)" />;
  if (fName.includes('İKTİSADİ') || fName.includes('İŞLETME') || fName.includes('YÖNETİM')) return <Briefcase size={48} strokeWidth={1.5} color="var(--color-icon-iktisat)" />;
  
  return <GraduationCap size={48} strokeWidth={1.5} color="var(--color-accent)" />;
};

const selectStyles = {
  control: (base, state) => ({
    ...base,
    borderRadius: 0,
    border: '2px solid var(--color-border)',
    boxShadow: state.isFocused ? '4px 4px 0 var(--color-accent)' : 'none',
    borderColor: state.isFocused ? 'var(--color-accent)' : 'var(--color-border)',
    '&:hover': {
      borderColor: state.isFocused ? 'var(--color-accent)' : 'var(--color-border)',
    },
    minHeight: '42px',
    background: 'var(--color-bg-main)',
    fontFamily: "'Inter', sans-serif",
    fontSize: '0.9rem',
    fontWeight: '500',
    transition: 'all 0.15s ease-out',
    transform: state.isFocused ? 'translate(-2px, -2px)' : 'none'
  }),
  menu: (base) => ({
    ...base,
    borderRadius: 0,
    border: '2px solid var(--color-border)',
    boxShadow: '6px 6px 0 var(--color-border)',
    marginTop: '4px',
    zIndex: 9999
  }),
  option: (base, state) => ({
    ...base,
    backgroundColor: state.isFocused ? 'var(--color-accent)' : 'transparent',
    color: state.isFocused ? 'white' : 'var(--color-text-main)',
    cursor: 'pointer',
    fontFamily: "'Inter', sans-serif",
    fontSize: '0.9rem',
    '&:active': {
      backgroundColor: 'var(--color-accent-hover)'
    }
  }),
  multiValue: (base) => ({
    ...base,
    backgroundColor: '#e5e5e5',
    borderRadius: 0,
    border: '1px solid var(--color-border)'
  }),
  multiValueLabel: (base) => ({
    ...base,
    color: 'var(--color-text-main)',
    fontWeight: 600,
    fontSize: '0.8rem'
  }),
  multiValueRemove: (base) => ({
    ...base,
    borderRadius: 0,
    '&:hover': {
      backgroundColor: 'var(--color-accent)',
      color: 'white',
    }
  })
};

function App() {
  const [programsData, setProgramsData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [isMobileFiltersOpen, setIsMobileFiltersOpen] = useState(false);

  // Pagination / Infinite Scroll
  const [displayCount, setDisplayCount] = useState(40);
  const observerTarget = useRef(null);

  const [searchTerm, setSearchTerm] = useState('');
  const [selectedUnis, setSelectedUnis] = useState([]);
  const [selectedFaculties, setSelectedFaculties] = useState([]);
  const [selectedPrograms, setSelectedPrograms] = useState([]);
  const [selectedPuanTypes, setSelectedPuanTypes] = useState([]);
  const [selectedUniTypes, setSelectedUniTypes] = useState([]);
  const [minPuan, setMinPuan] = useState('');
  const [maxPuan, setMaxPuan] = useState('');
  const [sortBy, setSortBy] = useState('puan-desc');

  useEffect(() => {
    const loadData = async () => {
      try {
        const response = await fetch('/en-kucuk-ve-en-buyuk-puanlar-tablo-4-rps0lq-18092428.xlsx');
        const arrayBuffer = await response.arrayBuffer();
        
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        const parsedData = [];
        
        for (let i = 4; i < rawData.length; i++) {
          const row = rawData[i];
          if (row && row[0] && String(row[0]).trim() !== '') {
            parsedData.push({
              programKodu: String(row[0] || '').trim(),
              universiteTuru: String(row[1] || '').trim(),
              universiteAdi: String(row[2] || '').trim(),
              fakulteAdi: String(row[3] || '').trim(),
              programAdi: String(row[4] || '').trim(),
              puanTuru: String(row[5] || '').trim(),
              genelKontenjan: row[6] || '-',
              enKucukPuan: row[8] || '--',
              enBuyukPuan: row[9] || '--'
            });
          }
        }
        
        setProgramsData(parsedData);
      } catch (error) {
        console.error("Excel dosyası okunurken hata oluştu:", error);
      } finally {
        setLoading(false);
      }
    };
    
    loadData();
  }, []);

  const uniOptions = useMemo(() => {
    let data = programsData;
    if (selectedUniTypes.length > 0) {
      data = data.filter(p => selectedUniTypes.some(t => t.value === p.universiteTuru));
    }
    return Array.from(new Set(data.map(p => p.universiteAdi).filter(Boolean))).sort().map(val => ({ value: val, label: val }));
  }, [programsData, selectedUniTypes]);

  const facOptions = useMemo(() => {
    let data = programsData;
    if (selectedUnis.length > 0) {
      data = data.filter(p => selectedUnis.some(u => u.value === p.universiteAdi));
    }
    return Array.from(new Set(data.map(p => p.fakulteAdi).filter(Boolean))).sort().map(val => ({ value: val, label: val }));
  }, [programsData, selectedUnis]);

  const progOptions = useMemo(() => {
    let data = programsData;
    if (selectedUnis.length > 0) {
      data = data.filter(p => selectedUnis.some(u => u.value === p.universiteAdi));
    }
    if (selectedFaculties.length > 0) {
      data = data.filter(p => selectedFaculties.some(f => f.value === p.fakulteAdi));
    }
    return Array.from(new Set(data.map(p => p.programAdi).filter(Boolean))).sort().map(val => ({ value: val, label: val }));
  }, [programsData, selectedUnis, selectedFaculties]);

  const puanOptions = useMemo(() => {
    let data = programsData;
    if (selectedUnis.length > 0) {
      data = data.filter(p => selectedUnis.some(u => u.value === p.universiteAdi));
    }
    return Array.from(new Set(data.map(p => p.puanTuru).filter(Boolean))).sort().map(val => ({ value: val, label: val }));
  }, [programsData, selectedUnis]);

  const uniTypeOptions = [
    { value: 'DEVLET', label: 'Devlet' },
    { value: 'VAKIF', label: 'Vakıf' }
  ];

  const trCustomFilter = (option, inputValue) => {
    return trToLower(option.label).includes(trToLower(inputValue));
  };

  const filteredPrograms = useMemo(() => {
    const result = programsData.filter(prog => {
      const matchSearch = searchTerm === '' || 
                          trToLower(prog.universiteAdi).includes(trToLower(searchTerm)) || 
                          trToLower(prog.programAdi).includes(trToLower(searchTerm));
                          
      const matchUniType = selectedUniTypes.length === 0 || selectedUniTypes.some(t => t.value === prog.universiteTuru);
      const matchPuanType = selectedPuanTypes.length === 0 || selectedPuanTypes.some(t => t.value === prog.puanTuru);
      const matchUni = selectedUnis.length === 0 || selectedUnis.some(u => u.value === prog.universiteAdi);
      const matchFaculty = selectedFaculties.length === 0 || selectedFaculties.some(f => f.value === prog.fakulteAdi);
      const matchProgram = selectedPrograms.length === 0 || selectedPrograms.some(p => p.value === prog.programAdi);

      const parseScore = (val) => {
        if (!val || val === '--') return NaN;
        if (typeof val === 'number') return val;
        return parseFloat(String(val).replace(',', '.'));
      };
      
      const score = parseScore(prog.enKucukPuan);
      const minFilter = minPuan === '' ? true : (!isNaN(score) && score >= parseFloat(minPuan));
      const maxFilter = maxPuan === '' ? true : (!isNaN(score) && score <= parseFloat(maxPuan));
      
      return matchSearch && matchUniType && matchPuanType && matchUni && matchFaculty && matchProgram && minFilter && maxFilter;
    });

    const parseScoreForSort = (val) => {
      if (!val || val === '--') return NaN;
      if (typeof val === 'number') return val;
      return parseFloat(String(val).replace(',', '.'));
    };

    result.sort((a, b) => {
      if (sortBy === 'puan-desc') {
        const scoreA = isNaN(parseScoreForSort(a.enKucukPuan)) ? -1 : parseScoreForSort(a.enKucukPuan);
        const scoreB = isNaN(parseScoreForSort(b.enKucukPuan)) ? -1 : parseScoreForSort(b.enKucukPuan);
        return scoreB - scoreA;
      }
      if (sortBy === 'puan-asc') {
        const scoreA = isNaN(parseScoreForSort(a.enKucukPuan)) ? Infinity : parseScoreForSort(a.enKucukPuan);
        const scoreB = isNaN(parseScoreForSort(b.enKucukPuan)) ? Infinity : parseScoreForSort(b.enKucukPuan);
        return scoreA - scoreB;
      }
      if (sortBy === 'uni-asc') {
        return trToLower(a.universiteAdi).localeCompare(trToLower(b.universiteAdi), 'tr');
      }
      if (sortBy === 'uni-desc') {
        return trToLower(b.universiteAdi).localeCompare(trToLower(a.universiteAdi), 'tr');
      }
      if (sortBy === 'fakulte-asc') {
        return trToLower(a.fakulteAdi).localeCompare(trToLower(b.fakulteAdi), 'tr');
      }
      if (sortBy === 'fakulte-desc') {
        return trToLower(b.fakulteAdi).localeCompare(trToLower(a.fakulteAdi), 'tr');
      }
      if (sortBy === 'bolum-asc') {
        return trToLower(a.programAdi).localeCompare(trToLower(b.programAdi), 'tr');
      }
      if (sortBy === 'bolum-desc') {
        return trToLower(b.programAdi).localeCompare(trToLower(a.programAdi), 'tr');
      }
      return 0;
    });

    return result;
  }, [searchTerm, selectedUniTypes, selectedPuanTypes, selectedUnis, selectedFaculties, selectedPrograms, minPuan, maxPuan, programsData, sortBy]);

  // Reset display count when filters change
  useEffect(() => {
    setDisplayCount(40);
  }, [filteredPrograms]);

  // Infinite Scroll Observer
  useEffect(() => {
    const observer = new IntersectionObserver(
      entries => {
        if (entries[0].isIntersecting) {
          setDisplayCount(prev => prev + 40);
        }
      },
      { threshold: 1.0 }
    );
    
    if (observerTarget.current) {
      observer.observe(observerTarget.current);
    }
    
    return () => {
      if (observerTarget.current) {
        observer.unobserve(observerTarget.current);
      }
    };
  }, [observerTarget]);

  if (loading) {
    return (
      <div className="modern-loader-container">
        <div className="brutalist-spinner"></div>
      </div>
    );
  }

  const displayedPrograms = filteredPrograms.slice(0, displayCount);

  const resetFilters = () => {
    setSearchTerm('');
    setSelectedUnis([]);
    setSelectedFaculties([]);
    setSelectedPrograms([]);
    setSelectedPuanTypes([]);
    setSelectedUniTypes([]);
    setMinPuan('');
    setMaxPuan('');
  };

  return (
    <div className="container">
      
      {/* Mobile Toggle Button */}
      <button 
        className="mobile-filter-toggle"
        onClick={() => setIsMobileFiltersOpen(!isMobileFiltersOpen)}
      >
        <Filter size={20} />
        {isMobileFiltersOpen ? "Filtreleri Gizle" : "Filtreleri Göster"}
      </button>

      {/* Sidebar Filters */}
      <aside className={`sidebar ${isMobileFiltersOpen ? 'mobile-open' : ''}`}>
        <h2>
          <Search size={24} style={{ marginRight: '8px', verticalAlign: 'middle' }}/>
          Filtreler
        </h2>
        
        <div className="filter-group">
          <label>Üniversite Türü</label>
          <Select
            isMulti
            options={uniTypeOptions}
            value={selectedUniTypes}
            onChange={setSelectedUniTypes}
            placeholder="Tümü"
            styles={selectStyles}
            filterOption={trCustomFilter}
          />
        </div>

        <div className="filter-group">
          <label htmlFor="search">Kelime ile Ara</label>
          <input 
            type="text" 
            id="search"
            placeholder="Üniversite veya bölüm adı..." 
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>

        <div className="filter-group">
          <label>Üniversiteler</label>
          <Select
            isMulti
            options={uniOptions}
            value={selectedUnis}
            onChange={setSelectedUnis}
            placeholder="Ara ve seç..."
            styles={selectStyles}
            noOptionsMessage={() => "Bulunamadı"}
            filterOption={trCustomFilter}
          />
        </div>

        <div className="filter-group">
          <label>Fakülteler</label>
          <Select
            isMulti
            options={facOptions}
            value={selectedFaculties}
            onChange={setSelectedFaculties}
            placeholder="Ara ve seç..."
            styles={selectStyles}
            noOptionsMessage={() => "Bulunamadı"}
            filterOption={trCustomFilter}
          />
        </div>

        <div className="filter-group">
          <label>Bölümler</label>
          <Select
            isMulti
            options={progOptions}
            value={selectedPrograms}
            onChange={setSelectedPrograms}
            placeholder="Ara ve seç..."
            styles={selectStyles}
            noOptionsMessage={() => "Bulunamadı"}
            filterOption={trCustomFilter}
          />
        </div>

        <div className="filter-group dual-input">
          <div>
            <label>Min Puan</label>
            <input type="number" placeholder="Örn: 300" value={minPuan} onChange={(e) => setMinPuan(e.target.value)} />
          </div>
          <div>
            <label>Max Puan</label>
            <input type="number" placeholder="Örn: 500" value={maxPuan} onChange={(e) => setMaxPuan(e.target.value)} />
          </div>
        </div>

        <div className="filter-group">
          <label>Puan Türü</label>
          <Select
            isMulti
            options={puanOptions}
            value={selectedPuanTypes}
            onChange={setSelectedPuanTypes}
            placeholder="Tümü"
            styles={selectStyles}
            filterOption={trCustomFilter}
          />
        </div>
        
        <button className="reset-btn" style={{ width: '100%', marginTop: '10px' }} onClick={resetFilters}>
          Tüm Filtreleri Temizle
        </button>

        {/* Sidebar Reklam Alanı */}
        <div style={{ marginTop: '20px' }}>
          <AdBanner dataAdSlot="SIDEBAR_SLOT_ID" />
        </div>
      </aside>

      {/* Main Content / Results */}
      <main className="main-content">
        <h1>2026 Lisans Programları</h1>
        
        <div className="results-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px' }}>
          <span>
            Toplam <strong>{filteredPrograms.length}</strong> program bulundu.
          </span>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <label style={{ fontSize: '0.8rem', fontWeight: 700, textTransform: 'uppercase', color: 'var(--color-text-muted)' }}>Sırala:</label>
            <select 
              value={sortBy} 
              onChange={(e) => setSortBy(e.target.value)}
              style={{
                padding: '6px 12px',
                border: '2px solid var(--color-border)',
                background: 'var(--color-bg-main)',
                fontFamily: "'Inter', sans-serif",
                fontWeight: 600,
                cursor: 'pointer',
                borderRadius: 0,
                outline: 'none'
              }}
            >
              <option value="puan-desc">Puan (Azalan)</option>
              <option value="puan-asc">Puan (Artan)</option>
              <option value="uni-asc">Üniversite (A-Z)</option>
              <option value="uni-desc">Üniversite (Z-A)</option>
              <option value="fakulte-asc">Fakülte (A-Z)</option>
              <option value="fakulte-desc">Fakülte (Z-A)</option>
              <option value="bolum-asc">Bölüm (A-Z)</option>
              <option value="bolum-desc">Bölüm (Z-A)</option>
            </select>
          </div>
        </div>

        <div className="results-grid">
          {displayedPrograms.length > 0 ? (
            <>
              {displayedPrograms.map((prog, index) => (
                <div key={`${prog.programKodu}-${index}`}>
                  {/* Her 20 sonuçta bir yatay reklam göster (0. index hariç) */}
                  {index > 0 && index % 20 === 0 && (
                    <AdBanner dataAdSlot="IN_FEED_SLOT_ID" dataAdFormat="fluid" />
                  )}
                  
                  <div className="program-card">
                    <div className="card-photo">
                      {getProgramIcon(prog.fakulteAdi, prog.programAdi)}
                    </div>
                    <div className="card-content">
                      <div className="card-header">
                        <div>
                          <div className="uni-name">{prog.universiteAdi} - {prog.fakulteAdi}</div>
                          <div className="prog-name">{prog.programAdi}</div>
                        </div>
                        <span className={`type-badge ${(prog.universiteTuru || '').toLocaleLowerCase('tr-TR')}`}>
                          {prog.universiteTuru}
                        </span>
                      </div>
                      
                      <div className="card-details">
                        <div className="detail-item">
                          <span className="detail-label">Puan Türü</span>
                          <span className={`detail-value score-badge ${(prog.puanTuru || '').toLocaleLowerCase('tr-TR')}`}>
                            {prog.puanTuru}
                          </span>
                        </div>
                        <div className="detail-item">
                          <span className="detail-label">Kontenjan</span>
                          <span className="detail-value">{prog.genelKontenjan}</span>
                        </div>
                        <div className="detail-item">
                          <span className="detail-label">Taban Puan</span>
                          <span className="detail-value highlight">
                            {typeof prog.enKucukPuan === 'number' ? prog.enKucukPuan.toFixed(5) : prog.enKucukPuan}
                          </span>
                        </div>
                        <div className="detail-item">
                          <span className="detail-label">Tavan Puan</span>
                          <span className="detail-value">
                            {typeof prog.enBuyukPuan === 'number' ? prog.enBuyukPuan.toFixed(5) : prog.enBuyukPuan}
                          </span>
                        </div>
                        <div className="detail-item">
                          <span className="detail-label">Program Kodu</span>
                          <span className="detail-value monospace">{prog.programKodu}</span>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              ))}
              
              {/* Infinite Scroll trigger element */}
              {displayedPrograms.length < filteredPrograms.length && (
                <div ref={observerTarget} style={{ height: '20px', width: '100%' }}></div>
              )}
            </>
          ) : (
            <div className="empty-state">
              <Search size={48} strokeWidth={1} />
              <p>Arama kriterlerinize uygun program bulunamadı.</p>
              <button className="reset-btn" onClick={resetFilters}>Filtreleri Temizle</button>
            </div>
          )}
        </div>
      </main>
    </div>
  );
}

export default App;

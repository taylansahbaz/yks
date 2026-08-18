import { useEffect, useRef } from 'react';

/**
 * Reusable Google AdSense Component for React
 * @param {string} dataAdSlot - AdSense reklam slot ID'si (Zorunlu)
 * @param {string} dataAdFormat - Ad formatı ('auto', 'rectangle', 'horizontal' vs)
 * @param {boolean} fullWidthResponsive - Tam genişlikte duyarlı (responsive) olacak mı?
 */
const AdBanner = ({ 
  dataAdSlot = "XXXXXXX", // TODO: Kendi Slot ID'nizi girin
  dataAdFormat = "auto", 
  fullWidthResponsive = true 
}) => {
  const adRef = useRef(null);

  useEffect(() => {
    try {
      // Sadece reklam henüz push edilmediyse push et
      // data-adsbygoogle-status="done" reklam yüklendiğinde eklenir
      if (adRef.current && !adRef.current.hasAttribute('data-adsbygoogle-status')) {
        (window.adsbygoogle = window.adsbygoogle || []).push({});
      }
    } catch (e) {
      console.error('AdSense hatası:', e);
    }
  }, []);

  return (
    <div className="ad-container" style={{ margin: 'var(--spacing-lg) 0', textAlign: 'center', overflow: 'hidden' }}>
      {/* 
        Gerçek reklam hesabınız onaylanana kadar geliştirme ortamında yer tutucu 
        olarak görünmesi için background eklendi.
      */}
      <ins
        ref={adRef}
        className="adsbygoogle"
        style={{ 
          display: 'block', 
          minHeight: '100px',
          background: 'var(--color-border-light)' // Placeholder rengi
        }}
        data-ad-client="ca-pub-XXXXXXXXXXXXXXXX" // TODO: Kendi Publisher Kimliğinizi (ca-pub) girin
        data-ad-slot={dataAdSlot}
        data-ad-format={dataAdFormat}
        data-full-width-responsive={fullWidthResponsive ? "true" : "false"}
      />
      <div style={{ fontSize: '0.7rem', color: 'var(--color-text-muted)', marginTop: '4px' }}>
        Reklam Alanı
      </div>
    </div>
  );
};

export default AdBanner;

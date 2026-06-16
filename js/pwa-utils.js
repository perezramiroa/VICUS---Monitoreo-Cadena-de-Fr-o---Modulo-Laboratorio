// ============================================
// UTILIDADES PWA PARA rsamio
// ============================================

class PWAUtils {
  
  // Registrar Service Worker con detección de actualizaciones
  static registerServiceWorker() {
    if ('serviceWorker' in navigator) {
      window.addEventListener('load', () => {
        navigator.serviceWorker.register('sw.js')
          .then(registration => {
            console.log('[PWA] Service Worker registrado:', registration.scope);

            // Buscar actualizaciones inmediatamente al cargar
            registration.update();

            // Intervalo para buscar actualizaciones cada 1 hora (opcional)
            setInterval(() => {
              registration.update();
            }, 1000 * 60 * 60);

            // Detectar cuando hay una nueva versión instalada y esperando
            registration.onupdatefound = () => {
              const installingWorker = registration.installing;
              installingWorker.onstatechange = () => {
                if (installingWorker.state === 'installed') {
                  if (navigator.serviceWorker.controller) {
                    console.log('[PWA] Nueva versión detectada. Aplicando...');
                    // Notificar al worker que puede tomar el control
                    installingWorker.postMessage({ type: 'SKIP_WAITING' });
                  }
                }
              };
            };
          })
          .catch(error => {
            console.error('[PWA] Error registrando SW:', error);
          });

        // Recargar automáticamente cuando el nuevo Service Worker toma el control
        let refreshing = false;
        navigator.serviceWorker.addEventListener('controllerchange', () => {
          if (!refreshing) {
            console.log('[PWA] El nuevo Service Worker ha tomado el control. Recargando...');
            window.location.reload();
            refreshing = true;
          }
        });
      });
    }
  }
  
  // Solicitar instalación como app
  static promptInstall() {
    let deferredPrompt;
    
    window.addEventListener('beforeinstallprompt', (e) => {
      // Prevenir el prompt automático
      e.preventDefault();
      deferredPrompt = e;
      
      // Mostrar botón de instalación
      this.showInstallButton(deferredPrompt);
    });
    
    window.addEventListener('appinstalled', () => {
      console.log('[PWA] App instalada exitosamente');
      // Ocultar botón de instalación
      const installBtn = document.getElementById('install-btn');
      if (installBtn) installBtn.style.display = 'none';
    });
  }
  
  // Mostrar botón de instalación
  static showInstallButton(deferredPrompt) {
    const installBtn = document.createElement('button');
    installBtn.id = 'install-btn';
    installBtn.className = 'btn-install-pwa';
    installBtn.innerHTML = '📱 <strong>Instalar App</strong>';
    installBtn.style.cssText = `
      position: fixed;
      bottom: 20px;
      right: 20px;
      background: linear-gradient(135deg, #4CAF50, #45a049);
      color: white;
      border: none;
      border-radius: 50px;
      padding: 12px 24px;
      font-size: 16px;
      font-weight: 600;
      box-shadow: 0 4px 15px rgba(0,0,0,0.3);
      cursor: pointer;
      z-index: 10000;
      transition: all 0.3s ease;
    `;
    
    installBtn.onmouseover = () => {
      installBtn.style.transform = 'scale(1.05)';
      installBtn.style.boxShadow = '0 6px 20px rgba(0,0,0,0.4)';
    };
    
    installBtn.onmouseout = () => {
      installBtn.style.transform = 'scale(1)';
    };
    
    installBtn.onclick = () => {
      deferredPrompt.prompt();
      deferredPrompt.userChoice.then((choiceResult) => {
        if (choiceResult.outcome === 'accepted') {
          console.log('[PWA] Usuario aceptó instalar');
        } else {
          console.log('[PWA] Usuario rechazó instalar');
        }
        installBtn.style.display = 'none';
      });
    };
    
    document.body.appendChild(installBtn);
  }
  
  // Solicitar permisos de notificación
  static async requestNotificationPermission() {
    if ('Notification' in window) {
      const permission = await Notification.requestPermission();
      
      if (permission === 'granted') {
        console.log('[PWA] Notificaciones permitidas');
      }
    }
  }
  
  // Verificar conexión offline
  static checkOnlineStatus() {
    window.addEventListener('online', () => {
      console.log('[PWA] Conexión restaurada');
    });
    
    window.addEventListener('offline', () => {
      console.log('[PWA] Sin conexión a internet');
    });
  }
  
  // Inicializar PWA
  static init() {
    this.registerServiceWorker();
    this.promptInstall();
    this.checkOnlineStatus();
    
    // Solicitar notificaciones después de 5 segundos
    setTimeout(() => {
      this.requestNotificationPermission();
    }, 5000);
    
    console.log('[PWA] rsamio PWA inicializado');
  }
}

// Inicializar automáticamente
document.addEventListener('DOMContentLoaded', () => {
  PWAUtils.init();
});
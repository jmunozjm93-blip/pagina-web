/* Pantalla de acceso del sitio (usuario y contraseña compartidos).
 *
 * AVISO: esto es una cortina, no una puerta con llave. El navegador descarga la
 * página completa antes de pedir la clave, así que alguien con conocimientos puede
 * saltarla. Sirve para que nadie entre por curiosidad o desde un buscador.
 * Para acceso real (correo @comercialdepor.cl) hay que poner el sitio detrás de
 * Cloudflare Access; está explicado en RESUMEN PROYECTO.md.
 *
 * Para cambiar la clave: calcular el SHA-256 de "usuario:contraseña" (usuario en
 * minúsculas) y reemplazar HASH. En PowerShell:
 *   $s="depor:NuevaClave"; $h=[System.Security.Cryptography.SHA256]::Create()
 *   ($h.ComputeHash([Text.Encoding]::UTF8.GetBytes($s))|%{$_.ToString("x2")}) -join ''
 */
(function () {
  'use strict';

  var HASH = '1b3dfbd511908a30fb9fa94a38f3b8ac1bbe94b02d4e7884a75625ed555e6e9d';
  var CLAVE = 'depor_acceso';
  var DIAS = 30;

  // la copia local (doble clic en tu propio PC) no pide nada
  if (location.protocol === 'file:') return;
  // sin crypto.subtle no se puede verificar; no dejamos el sitio trabado
  if (!(window.crypto && crypto.subtle)) return;

  function sesionVigente() {
    try {
      var v = JSON.parse(localStorage.getItem(CLAVE) || 'null');
      return !!(v && v.h === HASH && v.exp > Date.now());
    } catch (e) { return false; }
  }
  function guardarSesion() {
    try { localStorage.setItem(CLAVE, JSON.stringify({ h: HASH, exp: Date.now() + DIAS * 864e5 })); } catch (e) {}
  }
  async function sha256(txt) {
    var buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(txt));
    return [].map.call(new Uint8Array(buf), function (b) { return ('0' + b.toString(16)).slice(-2); }).join('');
  }

  if (sesionVigente()) return;

  var css = document.createElement('style');
  css.textContent = [
    '#acceso{position:fixed;inset:0;z-index:100000;background:#0b0b0b;display:flex;align-items:center;justify-content:center;padding:20px;',
    'font-family:"Helvetica Neue",Helvetica,Arial,system-ui,sans-serif;overflow:auto}',
    '#acceso .caja{width:100%;max-width:340px;text-align:center}',
    '#acceso h2{color:#fff;font-size:1.35rem;font-weight:800;letter-spacing:-.02em;margin:0 0 6px;text-transform:uppercase}',
    '#acceso p{color:#8a919b;font-size:.82rem;margin:0 0 22px;line-height:1.5}',
    '#acceso input{width:100%;background:#15181d;color:#fff;border:1px solid #2a3039;border-radius:8px;',
    'padding:12px 13px;font-size:16px;margin-bottom:10px;font-family:inherit}',
    '#acceso input:focus{outline:none;border-color:#ea580c}',
    '#acceso button{width:100%;background:#ea580c;color:#fff;border:0;border-radius:8px;padding:12px;',
    'font-weight:700;font-size:.95rem;cursor:pointer;font-family:inherit}',
    '#acceso button:disabled{opacity:.5;cursor:default}',
    '#acceso .err{color:#ef4444;font-size:.82rem;min-height:1.2em;margin-top:12px}'
  ].join('');
  document.head.appendChild(css);

  var box = document.createElement('div');
  box.id = 'acceso';
  box.innerHTML = '<form class="caja" autocomplete="on">' +
    '<h2>Sell Out · Grupo Depor</h2>' +
    '<p>Contenido interno. Ingresa con el usuario y la clave del equipo.</p>' +
    '<input id="acUsr" type="text" placeholder="Usuario" autocomplete="username" autocapitalize="off" autocorrect="off" spellcheck="false">' +
    '<input id="acPwd" type="password" placeholder="Contraseña" autocomplete="current-password">' +
    '<button id="acBtn" type="submit">Entrar</button>' +
    '<div class="err" id="acErr"></div></form>';

  // el script puede correr antes de que exista <body> (páginas sin esa etiqueta)
  if (document.body) document.body.appendChild(box);
  else document.addEventListener('DOMContentLoaded', function () { document.body.appendChild(box); });
  setTimeout(function () { var u = document.getElementById('acUsr'); if (u) u.focus(); }, 60);

  box.querySelector('form').addEventListener('submit', async function (e) {
    e.preventDefault();
    var btn = document.getElementById('acBtn'), err = document.getElementById('acErr');
    var usr = document.getElementById('acUsr').value.trim().toLowerCase();
    var pwd = document.getElementById('acPwd').value;
    if (!usr || !pwd) { err.textContent = 'Escribe usuario y contraseña.'; return; }
    btn.disabled = true; err.textContent = '';
    try {
      if (await sha256(usr + ':' + pwd) === HASH) {
        guardarSesion();
        box.remove(); css.remove();
        return;
      }
      err.textContent = 'Usuario o contraseña incorrectos.';
    } catch (ex) {
      err.textContent = 'No se pudo verificar: ' + ex.message;
    }
    btn.disabled = false;
    document.getElementById('acPwd').value = '';
    document.getElementById('acPwd').focus();
  });
})();

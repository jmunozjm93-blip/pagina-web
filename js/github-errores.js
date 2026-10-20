// Diagnósticos de GitHub sin incluir credenciales en mensajes o registros.
(function(global){
 function mensaje(r,j={}){
  const m=String(j.message||'');
  if(r.status===401)return 'El token no es válido, caducó o fue revocado. Revisa su estado en GitHub y vuelve a probar.';
  if(r.status===429 || ((r.status===403) && (r.headers.get('x-ratelimit-remaining')==='0'||r.headers.get('retry-after')||/rate limit|abuse/i.test(m))))return 'GitHub limitó temporalmente las solicitudes. Espera y vuelve a intentar; no necesitas cambiar el token por este aviso.';
  if(r.status===403 && r.headers.get('x-github-sso'))return 'GitHub requiere autorizar este token con el inicio de sesión de la organización (SSO).';
  if(r.status===403 && /resource not accessible|insufficient permission/i.test(m))return 'El token no tiene acceso suficiente. En GitHub, revisa que incluya pagina-web y el permiso Contents: Read and write. También comprueba si necesita aprobación del propietario.';
  if(r.status===403)return 'GitHub rechazó la operación (403). Puede deberse a permisos, políticas del repositorio o restricciones temporales. Revisa el acceso del token y las reglas de la rama; no es necesario crear otro token sin identificar la causa.';
  if(r.status===404)return 'GitHub no permite encontrar este recurso. Comprueba el repositorio y el acceso del token.';
  if(r.status>=500)return 'GitHub tiene un error temporal. Intenta nuevamente más tarde.';
  return `GitHub ${r.status}: ${m||r.statusText}`;
 }
 global.GitHubErrores={mensaje};
})(typeof window!=='undefined'?window:globalThis);

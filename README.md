# InfoChat (1999-2000)

## Contexto histórico

Aplicación de mensajería instantánea desarrollada, como parte de [Pdeinfo.com](https://pdeinfo.com), portal web de Piura, Perú, entre 1999-2000, orientado a servicios comunitarios para la región norte del país. El portal incluía sistema de anuncios clasificados, base de datos de currículos, listas de interés, postales digitales e InfoChat.

El desarrollo se realizó en contexto de recursos limitados: ciudad pequeña de provincia, sin comunidad técnica local establecida, sin capital de riesgo, conectividad precaria típica de Latinoamérica en esa época. El proyecto alcanzó tracción suficiente para generar oferta de adquisición por parte de operadores de Internet.

## Arquitectura técnica

### Stack tecnológico

**Backend:**

- Perl/CGI para servicios web stateless
- Archivos planos como almacenamiento (delimitadores `:` y `^`)
- Apache en hosting compartido

**Cliente chat:**

- Visual Basic 6 con controles Winsock (MSWINSCK.OCX)
- Microsoft Internet Transfer Control (MSINET.OCX) para peticiones HTTP
- Arquitectura peer-to-peer con discovery vía CGI

**Protocolo:**

- Puerto TCP 75 para conexiones directas entre clientes
- Protocolo aplicación de texto plano con comandos: `<c>` (chat mode), `<p>` (minipostal), `<s>` (confirmación), `<n>` (rechazo)

### Decisiones de diseño

**¿Por qué P2P en lugar de servidor centralizado?**

Razones pragmáticas bajo restricciones reales:

- Hosting compartido no permitía daemons persistentes en background
- Costos operativos de servidor dedicado inviables para proyecto autofinanciado
- Administración continua de procesos socket requería recursos no disponibles
- P2P eliminaba punto único de fallo y escalaba naturalmente con usuarios

**Rol de CGI:**

Scripts Perl actuaban como directory service mínimo:

- `addnick.pl`: Registro de presencia (nick:IP en archivo por fecha)
- `busnick.pl`: Resolución nick → IP del usuario conectado
- `regnick.pl`: Gestión de cuentas de usuario
- `encuentra.pl`: Búsqueda en base de datos usuarios

Los mensajes de chat viajaban directamente entre clientes sin pasar por servidor. CGI solo resolvía direcciones.

**Limitaciones conocidas:**

- Sin file locking: escrituras concurrentes pueden corromper archivos
- Passwords en texto plano en `registrados.bd`
- Validación inexistente: inyección trivial con delimiters en inputs
- Búsqueda O(n) lineal en archivos completos
- Requiere IP pública y puerto 75 entrante abierto

Estas limitaciones no impidieron operación funcional en el contexto objetivo: usuarios universitarios (UDEP) y empresariales con IPs públicas, comunidad geográficamente concentrada.

## Flujo de operación

1. Cliente A se registra/autentica vía HTTP con CGI
2. Cliente A inicia listener en puerto 75 local
3. Cada 10 minutos, cliente actualiza presencia: POST a `addnick.pl` con nick:IP actual
4. Cliente A desea chatear con usuario B
5. Cliente A consulta `busnick.pl?nick=B` → obtiene IP de B
6. Cliente A conecta directamente a `IP_B:75` vía Winsock TCP
7. Conexión persistente A↔B, mensajes directos sin intermediarios
8. Timer 50ms en UI procesa mensajes entrantes de variable global compartida

## Estructura del código

```
/cgi-bin/          # Scripts Perl CGI
  addnick.pl       # Registro presencia usuario
  busnick.pl       # Resolución nick → IP
  regnick.pl       # Registro nuevos usuarios
  encuentra.pl     # Búsqueda usuarios
  suscribe.pl      # Gestión listas de correo (SMTP integrado)
  pdeinfo.pl       # Contenido portal
  *.bd             # Archivos datos (registrados, postales, pdeinfo)

/src/              # Cliente VB6
  Finfo.frm        # Formulario principal: gestión contactos, Winsock listeners
  Fchat.frm        # Ventana de conversación individual
  Flogin.frm       # Autenticación
  Fregistrar.frm   # Registro usuarios nuevos
  power.bas        # Módulo utilidades y variables globales
  hlov.bas         # Obtención IP local
```

## IdSw

Este código representa caso de estudio auténtico para ingeniería de software:

**Limitaciones que motivan soluciones modernas:**

- Archivos planos sin transacciones → necesidad de ACID y bases de datos relacionales
- P2P con NAT/firewall → necesidad de servidores relay y STUN/TURN
- Polling y estateless → necesidad de WebSockets y conexiones persistentes
- Parsing manual sin validación → necesidad de ORMs, tipos estrictos, sanitización

**Decisiones de ingeniería bajo restricciones reales:**

- Trade-offs entre costos operativos y experiencia de usuario
- Arquitectura dictada por infraestructura disponible, no por patrones ideales
- Software funcional con imperfecciones técnicas > arquitectura perfecta sin usuarios

**Contexto comercial completo:**

- Desarrollo en mercado emergente con recursos limitados
- Tracción real demostrada (oferta de adquisición por Terra)
- Decisiones de negocio (rechazar oferta por control y timing) con consecuencias no obvias

## 2Think

El éxito comercial del proyecto no derivó de arquitectura técnicamente perfecta sino de resolver necesidad real de comunicación en comunidad específica con infraestructura apropiada para esa comunidad.

Las limitaciones técnicas evidentes (ausencia de file locking, passwords texto plano, P2P dependiente de IPs públicas) no impidieron adopción real ni valor comercial demostrable. Esto subraya que ingeniería de software es resolver problemas concretos con restricciones reales, no implementar patrones de libro de texto.

La preservación del código 25 años después permite análisis completo: desde requisitos originales y decisiones de diseño hasta consecuencias a largo plazo. Material didáctico con autenticidad imposible de replicar con ejemplos académicos sanitizados.

## Estado Actual

El portal https://www.pdeinfo.com permanece activo con propósitos históricos y educativos. El código fuente se preserva como documento técnico y testimonio del desarrollo web en Latinoamérica a finales de los años 90.

## Metadata Técnico

- **Período de desarrollo:** 1999-2000
- **Ubicación:** Piura, Perú
- **Afiliación:** Universidad de Piura (UDEP)
- **Lenguajes:** Perl, Visual Basic 6
- **Plataforma:** Windows 95/98/ME, Apache/Linux
- **Protocolo:** TCP directo (P2P), HTTP para directory service
- **Estado:** Código histórico preservado, portal activo solo con fines demostrativos
```

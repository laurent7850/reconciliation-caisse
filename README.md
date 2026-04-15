# Réconciliation Caisses — Saint Kilda

App web locale (Vite + React + TypeScript) pour réconcilier les rapports Z du caissier avec le récap annuel `SAINT KILDA - CAISSES 2026.xlsx`.

## Fonctionnement

1. **Charger le récap annuel** — le fichier Excel avec un onglet par mois.
2. **Charger les exports caisse** — plusieurs paires de fichiers :
   - `ReportZStats_1_<N>.xlsx` → N° Z, date, CA TTC, ventilation TVA
   - `CA_1_<N>.xlsx` → détail paiements (Cash, Carte banque, Virement)
   - Les autres exports (DetailSales, CAByItemLevel, CAbyGlobalUser) sont ignorés.
3. **Calculer la réconciliation** — l'app détecte le mois/jour à partir de la date d'ouverture du Z et propose le remplissage de la ligne correspondante.
4. **Télécharger** — un nouveau XLSX `...reconcilie-YYYY-MM-DD.xlsx` est généré. **Les cellules déjà remplies ne sont jamais écrasées**.

## Mapping

| Cellule récap | Source |
|---|---|
| `Z N°` | `ReportZStats.Rapport Z` |
| `TOTAL TVAC` | `ReportZStats.CA TTC` |
| `TOTAL 21% TVAC` | `ReportZStats.TTC 21%` |
| `TOTAL 12% TVAC` | `ReportZStats.TTC 12%` |
| `TOTAL 6% TVAC` | `ReportZStats.TTC 6%` |
| `Paiements cartes` | `CA.Carte banque` |
| `Virement client` | `CA.Virement bancaire` |
| `CASH` | `CA.Cash` |

## Conflits

- **Nouveau** (vert) : la ligne cible est vide → l'app remplit.
- **Déjà rempli (cohérent)** (orange) : la valeur existante = la valeur Z (±0,005 €) → rien à faire.
- **Conflit** (rouge) : la valeur existante diffère. L'app **ne touche pas** la cellule et signale le conflit dans le tableau + liste détaillée.

## Dev

```bash
npm install
npm run dev        # http://localhost:5173
npm run build      # build statique dans dist/
npm run preview    # sert le build local
```

## Déploiement VPS Hostinger

Build statique → copier `dist/` dans un vhost ou un container. Pas de backend.

```bash
npm run build
# upload dist/ vers /var/www/reconciliation/
```

Ou via Docker :

```dockerfile
FROM nginx:alpine
COPY dist/ /usr/share/nginx/html/
```

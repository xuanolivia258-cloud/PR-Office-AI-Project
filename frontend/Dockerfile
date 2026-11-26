FROM nginx:1.27-alpine

COPY frontend/nginx.conf /etc/nginx/conf.d/default.conf
COPY index.html script.js styles.css tracker.css xlsx.full.min.js /usr/share/nginx/html/

EXPOSE 80

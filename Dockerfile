FROM php:8.3-apache

ENV DEBIAN_FRONTEND=noninteractive
ENV PATH="/opt/venv/bin:${PATH}"

RUN apt-get update && apt-get install -y --no-install-recommends \
    git \
    unzip \
    zip \
    python3 \
    python3-venv \
    python3-pip \
    libzip-dev \
    libpng-dev \
    libjpeg62-turbo-dev \
    libfreetype6-dev \
    libonig-dev \
    && docker-php-ext-configure gd --with-freetype --with-jpeg \
    && docker-php-ext-install -j$(nproc) gd zip mbstring \
    && rm -rf /var/lib/apt/lists/*

COPY --from=composer:2 /usr/bin/composer /usr/bin/composer

RUN python3 -m venv /opt/venv \
    && pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir msoffcrypto-tool

WORKDIR /var/www/html

COPY . .

RUN if [ -f composer.json ]; then composer install --no-dev --prefer-dist --no-interaction --optimize-autoloader; else composer require phpoffice/phpspreadsheet:^5.7 --no-interaction --no-progress; fi

COPY docker/php.ini /usr/local/etc/php/conf.d/email-sorting.ini
COPY docker/render-start.sh /usr/local/bin/render-start

RUN chown -R www-data:www-data /var/www/html \
    && chmod +x /usr/local/bin/render-start \
    && chmod 755 /var/www/html/decrypt.py \
    && a2enmod headers \
    && php -m | grep -Ei "zip|mbstring|fileinfo|openssl|gd" \
    && python3 -c "import msoffcrypto; print('msoffcrypto OK')"

EXPOSE 10000

CMD ["render-start"]

# set base image (host OS)
FROM nginx:alpine

# set the working directory in the container
COPY . /usr/share/nginx/html

EXPOSE 80

FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build
WORKDIR /src
COPY . .

RUN dotnet restore "./halodownpatch/halodownpatch.csproj"
RUN dotnet publish "./halodownpatch/halodownpatch.csproj" -c Release -o /out

FROM nginx:alpine
COPY --from=build /out/wwwroot /usr/share/nginx/html
COPY nginx.conf /etc/nginx/conf.d/default.conf
EXPOSE 80

# Build stage
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# Copy solution and project files
COPY ReporteSAC.sln .
COPY src/Domain/Domain.csproj src/Domain/
COPY src/Application/Application.csproj src/Application/
COPY src/Infrastructure/Infrastructure.csproj src/Infrastructure/
COPY src/WebApp/WebApp.csproj src/WebApp/

# Restore packages
RUN dotnet restore

# Copy everything else and build
COPY . .
RUN dotnet publish src/WebApp/WebApp.csproj -c Release -o /app/publish

# Runtime stage
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS runtime
WORKDIR /app
COPY --from=build /app/publish .

# Render.com sets PORT env var at runtime
ENV ASPNETCORE_ENVIRONMENT=Production

EXPOSE 10000

# Use shell form so $PORT is expanded at runtime
CMD dotnet WebApp.dll --urls "http://+:${PORT:-10000}"

FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build
WORKDIR /src
COPY . .
RUN dotnet restore "./GoogleSheetsAndCsharp.csproj" --disable-parallel
RUN dotnet build "./GoogleSheetsAndCsharp.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "./GoogleSheetsAndCsharp.csproj" -c Release -o /app/publish --no-restore

FROM mcr.microsoft.com/dotnet/aspnet:6.0 AS base
WORKDIR /app
EXPOSE 5000

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "GoogleSheetsAndCsharp.dll"]
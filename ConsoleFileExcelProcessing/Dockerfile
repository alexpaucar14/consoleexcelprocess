FROM mcr.microsoft.com/dotnet/sdk:6.0
COPY . /app
RUN mkdir -p /app/data
RUN chmod o+x /app/data
WORKDIR /app
ENV TZ "America/Lima"

ENTRYPOINT ["dotnet", "ConsoleFileExcelProcessing.dll"]
using ExcelAndWordApi.Entities;
using ExcelAndWordApi.Persistence;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Identity;
using Microsoft.IdentityModel.Tokens;
using System.Text;

var MyAllowSpecificOrigins = "_myAllowSpecificOrigins";

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddCors(options =>
{
    options.AddPolicy(name: MyAllowSpecificOrigins,
                      policy =>
                      {
                          policy.AllowAnyOrigin();
                          policy.AllowAnyMethod();
                          policy.AllowAnyHeader();
                      });
});

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

// Configure the HTTP request pipeline.
    app.UseSwagger();
    app.UseSwaggerUI();

app.Use(async (context, next) =>
{
    string apiKey = context.Request.Headers["ApiKey"];

    if (string.IsNullOrWhiteSpace(apiKey) || app.Configuration.GetSection("ApiKey").Value != apiKey)
    {
        context.Response.StatusCode = 401; // Unauthorized
        await context.Response.WriteAsync("Invalid API Key");
        return;
    }

    await next();
});

app.UseHttpsRedirection();

app.UseCors(MyAllowSpecificOrigins);

app.MapControllers();

app.Run();

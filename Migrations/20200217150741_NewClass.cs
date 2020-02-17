using Microsoft.EntityFrameworkCore.Migrations;

namespace Teste.Migrations
{
    public partial class NewClass : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_Seller_TesteImport_TesteImportId",
                table: "Seller");

            migrationBuilder.AlterColumn<int>(
                name: "TesteImportId",
                table: "Seller",
                nullable: true,
                oldClrType: typeof(int),
                oldType: "int");

            migrationBuilder.AddForeignKey(
                name: "FK_Seller_TesteImport_TesteImportId",
                table: "Seller",
                column: "TesteImportId",
                principalTable: "TesteImport",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_Seller_TesteImport_TesteImportId",
                table: "Seller");

            migrationBuilder.AlterColumn<int>(
                name: "TesteImportId",
                table: "Seller",
                type: "int",
                nullable: false,
                oldClrType: typeof(int),
                oldNullable: true);

            migrationBuilder.AddForeignKey(
                name: "FK_Seller_TesteImport_TesteImportId",
                table: "Seller",
                column: "TesteImportId",
                principalTable: "TesteImport",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);
        }
    }
}

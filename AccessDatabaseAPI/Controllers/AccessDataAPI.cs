using Microsoft.AspNetCore.Mvc;
using System.Data.OleDb;
using System.Collections.Generic;
namespace AccessDatabaseAPI.Controllers
{

	[Route("api/[controller]")]
	[ApiController]
	public class AccessDataController : ControllerBase
	{
		private readonly string _connectionString;

		public AccessDataController(IConfiguration configuration)
		{
			_connectionString = configuration.GetConnectionString("AccessDb");
		}

		[HttpGet("GetDataFromTable")]
		public IActionResult GetDataFromTable([FromQuery] string tableName)
		{
			if (string.IsNullOrEmpty(tableName))
			{
				return BadRequest("Table name is required.");
			}

			var data = new List<Dictionary<string, object>>();

			using (OleDbConnection connection = new OleDbConnection(_connectionString))
			{
				try
				{
					connection.Open();
					string query = $"SELECT * FROM {tableName}";
					OleDbCommand command = new OleDbCommand(query, connection);
					OleDbDataReader reader = command.ExecuteReader();

					while (reader.Read())
					{
						var row = new Dictionary<string, object>();
						for (int i = 0; i < reader.FieldCount; i++)
						{
							row.Add(reader.GetName(i), reader[i]);
						}

						data.Add(row);
					}
				}
				catch (Exception ex)
				{
					return StatusCode(500, "Internal server error: " + ex.Message);
				}
			}

			return Ok(data);
		}

		[HttpPost("CreateRecord")]
		public IActionResult CreateRecord([FromQuery] string tableName, [FromBody] Dictionary<string, object> recordData)
		{
			if (string.IsNullOrEmpty(tableName) || recordData == null || recordData.Count == 0)
			{
				return BadRequest("Table name and record data are required.");
			}

			var allowedTables = new List<string> { "AUTHDEVICE", "ACGroup", "acholiday", "ACTimeZones", "ACUnlockComb", "AlarmLog", "AttParam", "AuditedExc", "Biotemplate", "CHECKEXACT", "CHECKINOUT", "Company3", "DEPARTMENTS", "DeptUsedSchs", "EmOpLog", "EXCNOTES", "FaceTemp", "HOLIDAYS", "ImportOptions", "LeaveClass", "LeaveClass1", "Machines", "NUM_RUN", "NUM_RUN_DEIL", "ReportItem", "SchClass", "SECURITYDETAILS", "ServerLog", "SHIFT", "SystemLog", "TBKEY", "TBSMSALLOT", "TBSMSINFO", "TEMPLATE", "USER_OF_RUN" }; 
			if (!allowedTables.Contains(tableName))
			{
				return BadRequest("Invalid table name.");
			}

			var columnNames = string.Join(", ", recordData.Keys);
			var parameterNames = string.Join(", ", recordData.Keys.Select(k => "@" + k));

			var query = $"INSERT INTO {tableName} ({columnNames}) VALUES ({parameterNames})";

			using (OleDbConnection connection = new OleDbConnection(_connectionString))
			{
				try
				{
					connection.Open();
					OleDbCommand command = new OleDbCommand(query, connection);
					foreach (var kvp in recordData)
					{
						object value = kvp.Value;
						if (value is DateTime)
						{
							command.Parameters.AddWithValue($"@{kvp.Key}", (DateTime)value);
						}
						else if (value is bool)
						{
							command.Parameters.AddWithValue($"@{kvp.Key}", (bool)value ? 1 : 0);
						}
						else if (value is string)
						{
							command.Parameters.AddWithValue($"@{kvp.Key}", (string)value);
						}
						else if (value is int || value is long || value is double || value is decimal)
						{
							command.Parameters.AddWithValue($"@{kvp.Key}", value);
						}
						else
						{
							command.Parameters.AddWithValue($"@{kvp.Key}", value?.ToString() ?? string.Empty);
						}
					}

					int rowsAffected = command.ExecuteNonQuery();

					if (rowsAffected > 0)
					{
						return Ok("Record created successfully.");
					}
					else
					{
						return StatusCode(500, "Failed to create the record.");
					}
				}
				catch (Exception ex)
				{
					return StatusCode(500, "Internal server error: " + ex.Message);
				}
			}
		}

	}
}




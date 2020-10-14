using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Cloud.Firestore;
using Google.Cloud.Firestore.V1;
using Grpc.Auth;
using Grpc.Core;

namespace ConsoleApp2
{
	//для доступа используется сервисный аккаунт firestorenetclient, см. по ссылке ниже
	//https://console.cloud.google.com/iam-admin/serviceaccounts?project=fitness365-677dc&authuser=1
	//https://googleapis.github.io/google-cloud-dotnet/docs/Google.Cloud.Firestore/#installation
	class Program
	{
		static async Task Main(string[] args)
		{
			var jsonConfig = File.ReadAllText(@"fitness365-677dc-888fd74f823c.json");

			var dbBuilder = new FirestoreDbBuilder();
			dbBuilder.ProjectId = "fitness365-677dc";
			dbBuilder.JsonCredentials = jsonConfig;

			FirestoreDb db = dbBuilder.Build();

			CollectionReference collectionReference = db.Collection("AppErrors");

			Query query = collectionReference
				.WhereLessThan("Date", DateTime.Now.AddDays(-5).ToUniversalTime());
			QuerySnapshot querySnapshot = await query.GetSnapshotAsync();

			foreach (DocumentSnapshot document in querySnapshot.Documents)
			{
				await document.Reference.DeleteAsync();
			}

			//Console.ReadKey();
		}
	}
}

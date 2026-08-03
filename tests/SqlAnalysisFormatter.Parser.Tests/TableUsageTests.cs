using SqlAnalysisFormatter.Parser;

namespace SqlAnalysisFormatter.Parser.Tests;

[TestClass]
public sealed class TableUsageTests
{
    [TestMethod]
    public void Build_Select_CollectsPhysicalInputsAndSkipsCteAndDerivedNames()
    {
        const string sql = """
            WITH active_users AS (
                SELECT u.id
                FROM dbo.users AS u
                JOIN dbo.accounts AS a ON a.user_id = u.id
            )
            SELECT au.id
            FROM active_users AS au
            JOIN (
                SELECT o.user_id
                FROM sales.orders AS o
            ) AS recent_orders ON recent_orders.user_id = au.id;
            """;

        var plan = OutputSheetPlanBuilder.Build(sql, []);

        CollectionAssert.AreEqual(
            new[] { "users", "accounts", "orders" },
            plan.InputTableIds.ToArray());
        Assert.IsEmpty(plan.OutputTableIds);
    }

    [TestMethod]
    public void Build_SelectInto_SeparatesSourceAndTargetTables()
    {
        const string sql = """
            SELECT u.id
            INTO #wkuser
            FROM dbo.users AS u
            LEFT JOIN dbo.locations AS l ON l.user_id = u.id;
            """;

        var plan = OutputSheetPlanBuilder.Build(sql, []);

        CollectionAssert.AreEqual(
            new[] { "users", "locations" },
            plan.InputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "#wkuser" }, plan.OutputTableIds.ToArray());
    }

    [TestMethod]
    public void Build_MultipleInserts_AggregatesSelectSourcesAndAllTargets()
    {
        const string sql = """
            INSERT INTO dbo.user_archive (id)
            SELECT u.id FROM dbo.users AS u
            UNION ALL
            SELECT d.id FROM dbo.deleted_users AS d;

            INSERT INTO dbo.audit_log (message)
            VALUES ('completed');
            """;

        var plan = OutputSheetPlanBuilder.Build(sql, []);

        CollectionAssert.AreEqual(
            new[] { "users", "deleted_users" },
            plan.InputTableIds.ToArray());
        CollectionAssert.AreEqual(
            new[] { "user_archive", "audit_log" },
            plan.OutputTableIds.ToArray());
    }

    [TestMethod]
    public void Build_InsertValues_DoesNotClassifyScalarSubqueryAsInputTable()
    {
        const string sql = """
            INSERT INTO dbo.audit_log (message)
            VALUES ((SELECT TOP (1) u.name FROM dbo.users AS u));
            """;

        var plan = OutputSheetPlanBuilder.Build(sql, []);

        Assert.IsEmpty(plan.InputTableIds);
        CollectionAssert.AreEqual(new[] { "audit_log" }, plan.OutputTableIds.ToArray());
    }

    [TestMethod]
    public void Build_UpdateAndDelete_IncludeResolvedTargetInBothDirections()
    {
        const string updateSql = """
            UPDATE u
            SET name = l.name
            FROM dbo.users AS u
            JOIN dbo.locations AS l ON l.user_id = u.id;
            """;
        const string deleteSql = """
            DELETE u
            FROM dbo.users AS u
            JOIN dbo.suspended_users AS s ON s.id = u.id;
            """;

        var updatePlan = OutputSheetPlanBuilder.Build(updateSql, []);
        var deletePlan = OutputSheetPlanBuilder.Build(deleteSql, []);

        CollectionAssert.AreEqual(
            new[] { "users", "locations" },
            updatePlan.InputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "users" }, updatePlan.OutputTableIds.ToArray());
        CollectionAssert.AreEqual(
            new[] { "users", "suspended_users" },
            deletePlan.InputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "users" }, deletePlan.OutputTableIds.ToArray());
    }

    [TestMethod]
    public void Build_Update_ReadsTargetOnlyWhenWhereOrNewValueReferencesIt()
    {
        const string constantSql = "UPDATE dbo.users SET name = 'fixed';";
        const string targetValueSql = "UPDATE dbo.users SET name = name + '!';";
        const string otherTableValueSql = """
            UPDATE dbo.users
            SET name = (SELECT d.name FROM dbo.defaults AS d);
            """;

        var constantPlan = OutputSheetPlanBuilder.Build(constantSql, []);
        var targetValuePlan = OutputSheetPlanBuilder.Build(targetValueSql, []);
        var otherTableValuePlan = OutputSheetPlanBuilder.Build(otherTableValueSql, []);

        Assert.IsEmpty(constantPlan.InputTableIds);
        CollectionAssert.AreEqual(new[] { "users" }, constantPlan.OutputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "users" }, targetValuePlan.InputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "users" }, targetValuePlan.OutputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "defaults" }, otherTableValuePlan.InputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "users" }, otherTableValuePlan.OutputTableIds.ToArray());
    }

    [TestMethod]
    public void Build_Update_PreservesTargetReadWhenAnotherTableColumnAppearsLater()
    {
        const string targetFirstSql = """
            UPDATE dbo.users
            SET name = users.name + d.suffix
            FROM dbo.defaults AS d;
            """;
        const string targetLastSql = """
            UPDATE dbo.users
            SET name = d.prefix + users.name
            FROM dbo.defaults AS d;
            """;

        var targetFirstPlan = OutputSheetPlanBuilder.Build(targetFirstSql, []);
        var targetLastPlan = OutputSheetPlanBuilder.Build(targetLastSql, []);

        CollectionAssert.AreEqual(
            new[] { "users", "defaults" },
            targetFirstPlan.InputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "users" }, targetFirstPlan.OutputTableIds.ToArray());
        CollectionAssert.AreEqual(
            new[] { "users", "defaults" },
            targetLastPlan.InputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "users" }, targetLastPlan.OutputTableIds.ToArray());
    }

    [TestMethod]
    public void Build_DmlOutputInto_ClassifiesSecondaryTargetsAsOutputs()
    {
        const string insertSql = """
            INSERT INTO dbo.users (name)
            OUTPUT inserted.id INTO dbo.audit_log (user_id)
            VALUES ('new user');
            """;
        const string updateSql = """
            UPDATE dbo.users
            SET name = 'fixed'
            OUTPUT deleted.id INTO dbo.audit_log (user_id)
            WHERE id = 1;
            """;
        const string deleteSql = """
            DELETE FROM dbo.users
            OUTPUT deleted.id INTO dbo.audit_log (user_id)
            WHERE id = 1;
            """;

        var insertPlan = OutputSheetPlanBuilder.Build(insertSql, []);
        var updatePlan = OutputSheetPlanBuilder.Build(updateSql, []);
        var deletePlan = OutputSheetPlanBuilder.Build(deleteSql, []);

        Assert.IsEmpty(insertPlan.InputTableIds);
        CollectionAssert.AreEqual(
            new[] { "users", "audit_log" },
            insertPlan.OutputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "users" }, updatePlan.InputTableIds.ToArray());
        CollectionAssert.AreEqual(
            new[] { "users", "audit_log" },
            updatePlan.OutputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "users" }, deletePlan.InputTableIds.ToArray());
        CollectionAssert.AreEqual(
            new[] { "users", "audit_log" },
            deletePlan.OutputTableIds.ToArray());
    }

    [TestMethod]
    public void Build_UpdateAndDelete_ResolveCteAndDerivedTargetsToPhysicalTable()
    {
        const string updateCteSql = """
            WITH target_users AS (
                SELECT u.id, u.name FROM dbo.users AS u
            )
            UPDATE target_users SET name = 'fixed';
            """;
        const string updateDerivedSql = """
            UPDATE target_user
            SET name = 'fixed'
            FROM (SELECT u.id, u.name FROM dbo.users AS u) AS target_user;
            """;
        const string deleteDerivedSql = """
            DELETE target_user
            FROM (SELECT u.id FROM dbo.users AS u) AS target_user;
            """;
        const string deleteCteSql = """
            WITH target_users AS (
                SELECT u.id FROM dbo.users AS u
            )
            DELETE FROM target_users WHERE id > 0;
            """;

        var updateCtePlan = OutputSheetPlanBuilder.Build(updateCteSql, []);
        var updateDerivedPlan = OutputSheetPlanBuilder.Build(updateDerivedSql, []);
        var deleteDerivedPlan = OutputSheetPlanBuilder.Build(deleteDerivedSql, []);
        var deleteCtePlan = OutputSheetPlanBuilder.Build(deleteCteSql, []);

        CollectionAssert.AreEqual(new[] { "users" }, updateCtePlan.InputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "users" }, updateCtePlan.OutputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "users" }, updateDerivedPlan.InputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "users" }, updateDerivedPlan.OutputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "users" }, deleteDerivedPlan.InputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "users" }, deleteDerivedPlan.OutputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "users" }, deleteCtePlan.InputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "users" }, deleteCtePlan.OutputTableIds.ToArray());
    }

    [TestMethod]
    public void Build_CteAndDerivedTargets_IgnorePredicateSubqueryTablesWhenResolvingOutput()
    {
        const string updateCteSql = """
            WITH target_users AS (
                SELECT u.id, u.name
                FROM dbo.users AS u
                WHERE EXISTS (
                    SELECT 1 FROM dbo.defaults AS d WHERE d.id = u.id
                )
            )
            UPDATE target_users SET name = 'fixed';
            """;
        const string updateDerivedSql = """
            UPDATE target_user
            SET name = 'fixed'
            FROM (
                SELECT u.id, u.name
                FROM dbo.users AS u
                WHERE EXISTS (
                    SELECT 1 FROM dbo.defaults AS d WHERE d.id = u.id
                )
            ) AS target_user;
            """;

        var updateCtePlan = OutputSheetPlanBuilder.Build(updateCteSql, []);
        var updateDerivedPlan = OutputSheetPlanBuilder.Build(updateDerivedSql, []);

        CollectionAssert.AreEqual(
            new[] { "users", "defaults" },
            updateCtePlan.InputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "users" }, updateCtePlan.OutputTableIds.ToArray());
        CollectionAssert.AreEqual(
            new[] { "users", "defaults" },
            updateDerivedPlan.InputTableIds.ToArray());
        CollectionAssert.AreEqual(new[] { "users" }, updateDerivedPlan.OutputTableIds.ToArray());
    }

    [TestMethod]
    public void Build_UnsupportedLaterStatementSuppressesAllTableUsage()
    {
        const string sql = """
            SELECT u.id FROM dbo.users AS u;
            INSERT INTO dbo.audit_log DEFAULT VALUES;
            """;

        var plan = OutputSheetPlanBuilder.Build(sql, []);

        Assert.IsEmpty(plan.InputTableIds);
        Assert.IsEmpty(plan.OutputTableIds);
    }

    [TestMethod]
    public void Build_FallbackHasNoTableUsage()
    {
        var plan = OutputSheetPlanBuilder.Build("SELECT FROM", []);

        Assert.IsTrue(plan.IsFallback);
        Assert.IsEmpty(plan.InputTableIds);
        Assert.IsEmpty(plan.OutputTableIds);
    }
}

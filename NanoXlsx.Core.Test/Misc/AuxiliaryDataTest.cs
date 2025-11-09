using Xunit;

namespace NanoXLSX.Core.Test.Misc
{
    public class AuxiliaryDataTest
    {
        [Theory(DisplayName = "Test of SetData and GetData with integer valueId")]
        [InlineData("plugin1", 0, "test_value")]
        [InlineData("plugin2", 42, 12345)]
        [InlineData("plugin3", 999, true)]
        public void SetDataGetDataWithIntValueIdTest(string pluginId, int valueId, object value)
        {
            var data = new AuxiliaryData();
            data.SetData(pluginId, valueId, value);

            var result = data.GetData(pluginId, valueId);
            Assert.Equal(value, result);
        }

        [Theory(DisplayName = "Test of SetData and GetData with string valueId")]
        [InlineData("plugin1", "key1", "test_value")]
        [InlineData("plugin2", "A1", 12345)]
        [InlineData("plugin3", "cell_ref", false)]
        public void SetDataGetDataWithStringValueIdTest(string pluginId, string valueId, object value)
        {
            var data = new AuxiliaryData();
            data.SetData(pluginId, valueId, value);

            var result = data.GetData(pluginId, valueId);
            Assert.Equal(value, result);
        }

        [Theory(DisplayName = "Test of SetData and GetData with entityId and integer valueId")]
        [InlineData("plugin1", "entity1", 0, "test_value")]
        [InlineData("plugin2", "worksheet1", 42, 12345)]
        [InlineData("plugin3", "sheet_A", 999, true)]
        public void SetDataGetDataWithEntityIdAndIntValueIdTest(string pluginId, string entityId, int valueId, object value)
        {
            var data = new AuxiliaryData();
            data.SetData(pluginId, entityId, valueId, value);

            var result = data.GetData<object>(pluginId, entityId, valueId);
            Assert.Equal(value, result);
        }

        [Theory(DisplayName = "Test of SetData and GetData with entityId and string valueId")]
        [InlineData("plugin1", "entity1", "key1", "test_value")]
        [InlineData("plugin2", "worksheet1", "A1", 12345)]
        [InlineData("plugin3", "sheet_A", "cell_ref", false)]
        public void SetDataGetDataWithEntityIdAndStringValueIdTest(string pluginId, string entityId, string valueId, object value)
        {
            var data = new AuxiliaryData();
            data.SetData(pluginId, entityId, valueId, value);

            var result = data.GetData(pluginId, entityId, valueId);
            Assert.Equal(value, result);
        }

        [Theory(DisplayName = "Test of generic GetData<T> with integer valueId")]
        [InlineData("plugin1", 0, "test_value")]
        [InlineData("plugin2", 42, 12345)]
        [InlineData("plugin3", 999, true)]
        public void GetDataGenericWithIntValueIdTest(string pluginId, int valueId, object value)
        {
            var data = new AuxiliaryData();
            data.SetData(pluginId, valueId, value);

            if (value is string)
            {
                var result = data.GetData<string>(pluginId, valueId);
                Assert.Equal(value, result);
            }
            else if (value is int)
            {
                var result = data.GetData<int>(pluginId, valueId);
                Assert.Equal(value, result);
            }
            else if (value is bool)
            {
                var result = data.GetData<bool>(pluginId, valueId);
                Assert.Equal(value, result);
            }
        }

        [Theory(DisplayName = "Test of generic GetData<T> with string valueId")]
        [InlineData("plugin1", "key1", "test_value")]
        [InlineData("plugin2", "A1", 12345)]
        [InlineData("plugin3", "cell_ref", true)]
        public void GetDataGenericWithStringValueIdTest(string pluginId, string valueId, object value)
        {
            var data = new AuxiliaryData();
            data.SetData(pluginId, valueId, value);

            if (value is string)
            {
                var result = data.GetData<string>(pluginId, valueId);
                Assert.Equal(value, result);
            }
            else if (value is int)
            {
                var result = data.GetData<int>(pluginId, valueId);
                Assert.Equal(value, result);
            }
            else if (value is bool)
            {
                var result = data.GetData<bool>(pluginId, valueId);
                Assert.Equal(value, result);
            }
        }

        [Theory(DisplayName = "Test of generic GetData<T> with entityId and integer valueId")]
        [InlineData("plugin1", "entity1", 0, "test_value")]
        [InlineData("plugin2", "worksheet1", 42, 12345)]
        [InlineData("plugin3", "sheet_A", 999, true)]
        public void GetDataGenericWithEntityIdAndIntValueIdTest(string pluginId, string entityId, int valueId, object value)
        {
            var data = new AuxiliaryData();
            data.SetData(pluginId, entityId, valueId, value);

            if (value is string)
            {
                var result = data.GetData<string>(pluginId, entityId, valueId);
                Assert.Equal(value, result);
            }
            else if (value is int)
            {
                var result = data.GetData<int>(pluginId, entityId, valueId);
                Assert.Equal(value, result);
            }
            else if (value is bool)
            {
                var result = data.GetData<bool>(pluginId, entityId, valueId);
                Assert.Equal(value, result);
            }
        }

        [Theory(DisplayName = "Test of generic GetData<T> with entityId and string valueId")]
        [InlineData("plugin1", "entity1", "key1", "test_value")]
        [InlineData("plugin2", "worksheet1", "A1", 12345)]
        [InlineData("plugin3", "sheet_A", "cell_ref", false)]
        public void GetDataGenericWithEntityIdAndStringValueIdTest(string pluginId, string entityId, string valueId, object value)
        {
            var data = new AuxiliaryData();
            data.SetData(pluginId, entityId, valueId, value);

            if (value is string)
            {
                var result = data.GetData<string>(pluginId, entityId, valueId);
                Assert.Equal(value, result);
            }
            else if (value is int)
            {
                var result = data.GetData<int>(pluginId, entityId, valueId);
                Assert.Equal(value, result);
            }
            else if (value is bool)
            {
                var result = data.GetData<bool>(pluginId, entityId, valueId);
                Assert.Equal(value, result);
            }
        }

        [Fact(DisplayName = "Test of GetData returning null for non-existent data")]
        public void GetDataNonExistentTest()
        {
            var data = new AuxiliaryData();

            var result1 = data.GetData("plugin1", 0);
            Assert.Null(result1);

            var result2 = data.GetData("plugin1", "key1");
            Assert.Null(result2);

            var result3 = data.GetData<object>("plugin1", "entity1", 0);
            Assert.Null(result3);

            var result4 = data.GetData("plugin1", "entity1", "key1");
            Assert.Null(result4);
        }

        [Fact(DisplayName = "Test of generic GetData<T> returning default for non-existent data")]
        public void GetDataGenericNonExistentTest()
        {
            var data = new AuxiliaryData();

            var result1 = data.GetData<string>("plugin1", 0);
            Assert.Null(result1);

            var result2 = data.GetData<int>("plugin1", "key1");
            Assert.Equal(0, result2);

            var result3 = data.GetData<bool>("plugin1", "entity1", 0);
            Assert.False(result3);

            var result4 = data.GetData<string>("plugin1", "entity1", "key1");
            Assert.Null(result4);
        }

        [Fact(DisplayName = "Test of generic GetData<T> returning default for wrong type")]
        public void GetDataGenericWrongTypeTest()
        {
            var data = new AuxiliaryData();
            data.SetData("plugin1", 0, "string_value");

            var result = data.GetData<int>("plugin1", 0);
            Assert.Equal(0, result);
        }

        [Fact(DisplayName = "Test of SetData update behavior")]
        public void SetDataUpdateTest()
        {
            var data = new AuxiliaryData();
            data.SetData("plugin1", 0, "initial_value");

            var result1 = data.GetData("plugin1", 0);
            Assert.Equal("initial_value", result1);

            data.SetData("plugin1", 0, "updated_value");
            var result2 = data.GetData("plugin1", 0);
            Assert.Equal("updated_value", result2);
        }

        [Fact(DisplayName = "Test of GetDataList with default entity")]
        public void GetDataListDefaultEntityTest()
        {
            var data = new AuxiliaryData();
            data.SetData("plugin1", 0, "value1");
            data.SetData("plugin1", 1, "value2");
            data.SetData("plugin1", 2, "value3");

            var result = data.GetDataList<string>("plugin1");
            Assert.Equal(3, result.Count);
            Assert.Contains("value1", result);
            Assert.Contains("value2", result);
            Assert.Contains("value3", result);
        }

        [Fact(DisplayName = "Test of GetDataList with specific entity")]
        public void GetDataListSpecificEntityTest()
        {
            var data = new AuxiliaryData();
            data.SetData("plugin1", "entity1", 0, "value1");
            data.SetData("plugin1", "entity1", 1, "value2");
            data.SetData("plugin1", "entity2", 0, "value3");

            var result = data.GetDataList<string>("plugin1", "entity1");
            Assert.Equal(2, result.Count);
            Assert.Contains("value1", result);
            Assert.Contains("value2", result);
            Assert.DoesNotContain("value3", result);
        }

        [Fact(DisplayName = "Test of GetDataList with mixed types")]
        public void GetDataListMixedTypesTest()
        {
            var data = new AuxiliaryData();
            data.SetData("plugin1", 0, "string_value");
            data.SetData("plugin1", 1, 42);
            data.SetData("plugin1", 2, "another_string");

            var result = data.GetDataList<string>("plugin1");
            Assert.Equal(2, result.Count);
            Assert.Contains("string_value", result);
            Assert.Contains("another_string", result);
        }

        [Fact(DisplayName = "Test of GetDataList returning empty list for non-existent plugin")]
        public void GetDataListNonExistentPluginTest()
        {
            var data = new AuxiliaryData();

            var result = data.GetDataList<string>("nonexistent_plugin");
            Assert.NotNull(result);
            Assert.Empty(result);
        }

        [Fact(DisplayName = "Test of persistent data with ClearTemporaryData")]
        public void PersistentDataTest()
        {
            var data = new AuxiliaryData();
            data.SetData("plugin1", 0, "temporary_value", false);
            data.SetData("plugin1", 1, "persistent_value", true);

            data.ClearTemporaryData();

            var temp = data.GetData("plugin1", 0);
            Assert.Null(temp);

            var persistent = data.GetData("plugin1", 1);
            Assert.Equal("persistent_value", persistent);
        }

        [Fact(DisplayName = "Test of ClearTemporaryData with multiple plugins and entities")]
        public void ClearTemporaryDataComplexTest()
        {
            var data = new AuxiliaryData();
            data.SetData("plugin1", "entity1", 0, "temp1", false);
            data.SetData("plugin1", "entity1", 1, "persist1", true);
            data.SetData("plugin1", "entity2", 0, "temp2", false);
            data.SetData("plugin2", "entity1", 0, "temp3", false);
            data.SetData("plugin2", "entity1", 1, "persist2", true);

            data.ClearTemporaryData();

            Assert.Null(data.GetData<object>("plugin1", "entity1", 0));
            Assert.Equal("persist1", data.GetData<object>("plugin1", "entity1", 1));
            Assert.Null(data.GetData<object>("plugin1", "entity2", 0));
            Assert.Null(data.GetData<object>("plugin2", "entity1", 0));
            Assert.Equal("persist2", data.GetData<object>("plugin2", "entity1", 1));
        }

        [Fact(DisplayName = "Test of ClearTemporaryData removing empty structures")]
        public void ClearTemporaryDataRemovesEmptyStructuresTest()
        {
            var data = new AuxiliaryData();
            data.SetData("plugin1", "entity1", 0, "temp1", false);
            data.SetData("plugin1", "entity2", 0, "temp2", false);

            data.ClearTemporaryData();

            // After clearing, getting data list should return empty
            var result1 = data.GetDataList<string>("plugin1", "entity1");
            Assert.Empty(result1);

            var result2 = data.GetDataList<string>("plugin1", "entity2");
            Assert.Empty(result2);
        }

        [Fact(DisplayName = "Test of ClearData")]
        public void ClearDataTest()
        {
            var data = new AuxiliaryData();
            data.SetData("plugin1", 0, "value1", false);
            data.SetData("plugin1", 1, "value2", true);
            data.SetData("plugin2", "entity1", 0, "value3", false);

            data.ClearData();

            Assert.Null(data.GetData("plugin1", 0));
            Assert.Null(data.GetData("plugin1", 1));
            Assert.Null(data.GetData<object>("plugin2", "entity1", 0));

            var result = data.GetDataList<string>("plugin1");
            Assert.Empty(result);
        }

        [Fact(DisplayName = "Test of data isolation between different plugins")]
        public void DataIsolationBetweenPluginsTest()
        {
            var data = new AuxiliaryData();
            data.SetData("plugin1", 0, "value1");
            data.SetData("plugin2", 0, "value2");

            var result1 = data.GetData<string>("plugin1", 0);
            var result2 = data.GetData<string>("plugin2", 0);

            Assert.Equal("value1", result1);
            Assert.Equal("value2", result2);
            Assert.NotEqual(result1, result2);
        }

        [Fact(DisplayName = "Test of data isolation between different entities")]
        public void DataIsolationBetweenEntitiesTest()
        {
            var data = new AuxiliaryData();
            data.SetData("plugin1", "entity1", 0, "value1");
            data.SetData("plugin1", "entity2", 0, "value2");

            var result1 = data.GetData<string>("plugin1", "entity1", 0);
            var result2 = data.GetData<string>("plugin1", "entity2", 0);

            Assert.Equal("value1", result1);
            Assert.Equal("value2", result2);
            Assert.NotEqual(result1, result2);
        }

        [Theory(DisplayName = "Test of GetData with integer entityId and string valueId")]
        [InlineData("plugin1", 0, "key1", "test_value")]
        [InlineData("plugin2", 42, "A1", 12345)]
        [InlineData("plugin3", 999, "cell_ref", false)]
        public void GetDataWithIntEntityIdAndStringValueIdTest(string pluginId, int entityId, string valueId, object value)
        {
            var data = new AuxiliaryData();
            data.SetData(pluginId, entityId, valueId, value);

            var result = data.GetData(pluginId, entityId, valueId);
            Assert.Equal(value, result);
        }
    }
}

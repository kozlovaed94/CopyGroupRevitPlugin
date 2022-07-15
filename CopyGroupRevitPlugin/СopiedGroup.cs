using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyGroupRevitPlugin
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class СopiedGroup : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uiDocument = commandData.Application.ActiveUIDocument;
                Document document = uiDocument.Document;

                Group selectedGroup = ObjectGroup.picktGroup(uiDocument, document);
                Room selectedRoomForPlacement = ObjectGroup.pickRoomToInsertTheGroup(uiDocument, document);
                XYZ insertionPoint = InsertionPoint.calculateInsertionPointWithOffsetGroupCenterRelativeRoomForPlacementCenter(document, selectedGroup, selectedRoomForPlacement);
                ObjectGroup.placeGroupWithOffsetFromRoomCenter(document, selectedGroup, insertionPoint);
                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    public static class ObjectGroup
    {
        public static Group picktGroup(UIDocument uiDocument, Document document)
        {
            GroupPickFilter groupPickFilter = new GroupPickFilter();
            Reference reference = uiDocument.Selection.PickObject(ObjectType.Element, groupPickFilter, "Выберете группу объектов");
            Group selectedGroup = document.GetElement(reference) as Group;
            return selectedGroup;
        }
        public static Room pickRoomToInsertTheGroup(UIDocument uiDocument, Document document)
        {
            XYZ pickedPointInRoomToInsertTheGroup = uiDocument.Selection.PickPoint("Выберете точку вставки");
            Room selectedRoom = ProjectRoom.GetRoomByPoint(document, pickedPointInRoomToInsertTheGroup);
            return selectedRoom;
        }        
        public static void placeGroupWithOffsetFromRoomCenter(Document document, Group selectedGroup, XYZ insertionPoint)
        {
            Transaction transaction = new Transaction(document);
            transaction.Start("Вставка группы объектов");
            document.Create.PlaceGroup(insertionPoint, selectedGroup.GroupType);
            transaction.Commit();
        }
    }

    public static class InsertionPoint
    {       
        public static XYZ calculateInsertionPointWithOffsetGroupCenterRelativeRoomForPlacementCenter(Document document, Group group, Room room)
        {
            XYZ roomCenter = ElementCenter.GetElementCenter(room);
            XYZ offset = calculateOffsetGroupCenterRelativeGroupPlacementRoomCenter(document, group);
            XYZ insertionGroupPoint = roomCenter + offset;
            XYZ insertionPoint = new XYZ(insertionGroupPoint.X, insertionGroupPoint.Y, room.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET).AsDouble());
            return insertionPoint;
        }
        private static XYZ calculateOffsetGroupCenterRelativeGroupPlacementRoomCenter(Document document, Group group)
        {
            XYZ groupCenter = ElementCenter.GetElementCenter(group);
            Room groupPlacementRoom = ProjectRoom.GetRoomByPoint(document, groupCenter);
            XYZ groupPlacementRoomCenter = ElementCenter.GetElementCenter(groupPlacementRoom);
            XYZ offset = groupCenter - groupPlacementRoomCenter;
            return offset;
        }
    }

    public static class ElementCenter
    {
        public static XYZ GetElementCenter(Element element)
        {
            BoundingBoxXYZ bounding = element.get_BoundingBox(null);
            return (bounding.Max + bounding.Min) / 2;
        }
    }

    public class GroupPickFilter : ISelectionFilter
    {
        public bool AllowElement(Element element)
        {
            if (element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_IOSModelGroups) return true;
            else return false;
        }
        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }

    public static class ProjectRoom
    {
        public static Room GetRoomByPoint(Document document, XYZ point)
        {
            FilteredElementCollector roomCollection = new FilteredElementCollector(document)
                                                                  .OfCategory(BuiltInCategory.OST_Rooms);
            foreach (Room room in roomCollection)
            {
                if (room != null)
                {
                    if (room.IsPointInRoom(point)) return room;
                }
            }
            return null;
        }
    }
}

import cv2


def update_center_points(data, dic_center_points):
    '''
    更新坐标
    '''
    for row in data:
        x1, y1, x2, y2, cls_name, conf, obj_id = row[:7]

        # 计算中心点坐标
        center_x = int((x1 + x2) / 2)
        center_y = int((y1 + y2) / 2)

        # 更新字典
        if obj_id in dic_center_points:
            # 判断列表长度是否超过30
            if len(dic_center_points[obj_id]) >= 30:
                dic_center_points[obj_id].pop(0)
            dic_center_points[obj_id].append((center_x, center_y))
        else:
            dic_center_points[obj_id] = [(center_x, center_y)]

    return dic_center_points




def res2OCres(results):
    lst_res = []
    if results is None:
        return lst_res
    for res in results.tolist():
        box = res[:4]
        conf = res[-2]
        cls = res[-1]
        lst_res.append([cls, conf, box])

    return list(lst_res)




def compute_color_for_labels(label):
    palette = (2 ** 11 - 1, 2 ** 15 - 1, 2 ** 20 - 1)
    """
    Simple function that adds fixed color depending on the class
    """
    color = [int((p * (label ** 2 - label + 1)) % 255) for p in palette]
    return tuple(color)


def draw_info(frame, results):
    for i, bbox in enumerate(results):
        cls_name = bbox[0]
        conf = bbox[1]
        box = bbox[2]

        color = compute_color_for_labels(i)
        # 彩框 yolov5
        cv2.rectangle(frame, (int(box[0]), int(box[1])), (int(box[2]), int(box[3])), color, 2)
        # cls_name
        cv2.putText(frame, str(cls_name), (int(box[0] + 20), int(box[1] + 50)), 0, 5e-3 * 80, color, 1)

    return frame




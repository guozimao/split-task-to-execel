package beans;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.PictureData;

/**
 * 存储图片数据，导出准备的model
 *
 * */
public class MyPicture {
    //图片生成的位置
    private ClientAnchor clientAnchor;

    //图片的数据
    private PictureData pictureData;

    public ClientAnchor getClientAnchor() {
        return clientAnchor;
    }

    public void setClientAnchor(ClientAnchor clientAnchor) {
        this.clientAnchor = clientAnchor;
    }

    public PictureData getPictureData() {
        return pictureData;
    }

    public void setPictureData(PictureData pictureData) {
        this.pictureData = pictureData;
    }
}

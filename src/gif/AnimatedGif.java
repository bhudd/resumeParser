package gif;

import java.awt.image.BufferedImage;
import java.io.InputStream;

import javafx.animation.Interpolator;
import javafx.animation.Transition;
import javafx.embed.swing.SwingFXUtils;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.image.WritableImage;
import javafx.util.Duration;

public class AnimatedGif extends Transition {

	private ImageView imageView;
	private int lastIndex;
	
	private Image[] sequence;
	
	public AnimatedGif(ImageView view, InputStream gifImage, Duration duration)
	{
		this.imageView = view;
		GifDecoder d = new GifDecoder();
		d.read(gifImage);
		sequence = new Image[d.getFrameCount()];
		for(int i=0; i < d.getFrameCount(); i++)
		{
			WritableImage wimg = null;
			BufferedImage bimg = d.getFrame(i);
			sequence[i] = SwingFXUtils.toFXImage(bimg, wimg);
		}
		
		setCycleDuration(duration);
		setInterpolator(Interpolator.LINEAR);
	}

	@Override
	protected void interpolate(double k) {
		final int index = Math.min((int) Math.floor(k * sequence.length), sequence.length - 1);
        if (index != lastIndex) {
            imageView.setImage(sequence[index]);
            lastIndex = index;
        }
	}
}
